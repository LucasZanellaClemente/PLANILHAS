"""Camada de interface: leitura de entrada do usuário e handlers do menu.

Este módulo concentra toda a interação com o terminal (perguntas, prévias,
confirmações) e delega as regras de negócio ao módulo ``operacoes``, a
importação de arquivos ao módulo ``carregador``, os metadados ao módulo
``catalogo`` e a geração do relatório final ao módulo ``exportador``. Nenhuma
lógica de transformação de dados deve viver aqui além da orquestração da
conversa com o usuário.
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd

import carregador
import catalogo
import exportador
import operacoes
from estado import Dataset, HistoricoEntry, Sessao
from utils import ErroExpressaoInsegura, ErroOperacao, avaliar_condicao_segura, avaliar_expressao_segura, truncar_texto

logger = logging.getLogger(__name__)

try:
    from tabulate import tabulate

    _TEM_TABULATE = True
except ImportError:  # a biblioteca é opcional; há um modo de exibição alternativo
    _TEM_TABULATE = False


# ---------------------------------------------------------------------------
# Helpers de entrada e exibição
# ---------------------------------------------------------------------------


def ler_texto(mensagem: str, obrigatorio: bool = True, padrao: Optional[str] = None) -> str:
    """Lê uma string do usuário, repetindo a pergunta enquanto o campo obrigatório estiver vazio.

    Propositalmente não trata ``EOFError``: se a entrada padrão terminar
    inesperadamente (ex.: pipe fechado), o erro sobe até o laço principal em
    ``main.py``, que encerra a aplicação de forma controlada em vez de cair
    em um laço infinito pedindo entradas que nunca chegarão.
    """
    while True:
        bruto = input(mensagem).strip()
        if not bruto and padrao is not None:
            return padrao
        if not bruto and obrigatorio:
            print("Este campo é obrigatório. Tente novamente.")
            continue
        return bruto


def ler_inteiro(mensagem: str, minimo: Optional[int] = None, maximo: Optional[int] = None, padrao: Optional[int] = None) -> int:
    """Lê um número inteiro do usuário, validando limites opcionais."""
    while True:
        bruto = input(mensagem).strip()
        if not bruto and padrao is not None:
            return padrao
        try:
            valor = int(bruto)
        except ValueError:
            print("Digite um número inteiro válido.")
            continue
        if minimo is not None and valor < minimo:
            print(f"O valor deve ser maior ou igual a {minimo}.")
            continue
        if maximo is not None and valor > maximo:
            print(f"O valor deve ser menor ou igual a {maximo}.")
            continue
        return valor


def ler_float(mensagem: str, padrao: Optional[float] = None) -> float:
    """Lê um número decimal do usuário, aceitando vírgula ou ponto."""
    while True:
        bruto = input(mensagem).strip().replace(",", ".")
        if not bruto and padrao is not None:
            return padrao
        try:
            return float(bruto)
        except ValueError:
            print("Digite um número válido.")


def confirmar(mensagem: str, padrao: bool = False) -> bool:
    """Pede uma confirmação sim/não ao usuário."""
    sufixo = "[S/n]" if padrao else "[s/N]"
    bruto = input(f"{mensagem} {sufixo}: ").strip().lower()
    if not bruto:
        return padrao
    return bruto in ("s", "sim", "y", "yes")


def _imprimir_tabela(linhas: list[dict[str, Any]]) -> None:
    """Imprime uma lista de dicionários como tabela, usando tabulate quando disponível."""
    if not linhas:
        print("(nenhum dado)")
        return
    if _TEM_TABULATE:
        print(tabulate(linhas, headers="keys", tablefmt="simple", showindex=False))
    else:
        print(pd.DataFrame(linhas).to_string(index=False))


def exibir_previa(df: pd.DataFrame, linhas: int = 10) -> None:
    """Exibe uma prévia das primeiras linhas de um DataFrame."""
    if df is None or df.empty:
        print("(sem dados para exibir)")
        return
    with pd.option_context("display.max_columns", 20, "display.width", 160, "display.max_colwidth", 30):
        print(df.head(linhas).to_string(index=False))
    if len(df) > linhas:
        print(f"... ({len(df) - linhas} linha(s) adicionais não exibidas)")


def listar_colunas_numeradas(df: pd.DataFrame) -> None:
    """Imprime a lista de colunas do DataFrame numeradas para seleção."""
    for indice, coluna in enumerate(df.columns, start=1):
        print(f"  [{indice}] {coluna}")


def selecionar_coluna(df: pd.DataFrame, mensagem: str = "Escolha a coluna pelo número", permitir_cancelar: bool = False) -> Optional[str]:
    """Pede ao usuário para escolher uma única coluna pelo número exibido."""
    listar_colunas_numeradas(df)
    minimo = 0 if permitir_cancelar else 1
    sufixo = " (0 para cancelar)" if permitir_cancelar else ""
    escolha = ler_inteiro(f"{mensagem}{sufixo}: ", minimo=minimo, maximo=len(df.columns))
    if permitir_cancelar and escolha == 0:
        return None
    return str(df.columns[escolha - 1])


def selecionar_colunas_multiplas(df: pd.DataFrame, mensagem: str = "Escolha as colunas (números separados por vírgula)") -> list[str]:
    """Pede ao usuário para escolher uma ou mais colunas pelos números exibidos."""
    listar_colunas_numeradas(df)
    bruto = ler_texto(f"{mensagem}: ", obrigatorio=False, padrao="")
    colunas: list[str] = []
    for parte in bruto.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if not parte.isdigit():
            print(f"Entrada inválida ignorada: '{parte}'")
            continue
        indice = int(parte)
        if 1 <= indice <= len(df.columns):
            colunas.append(str(df.columns[indice - 1]))
        else:
            print(f"Índice fora do intervalo ignorado: {indice}")
    return colunas


def selecionar_dataset(sessao: Sessao, mensagem: str = "Escolha o dataset pelo número") -> Optional[Dataset]:
    """Lista os datasets carregados e pede ao usuário para escolher um pelo identificador."""
    datasets = sessao.listar_datasets()
    if not datasets:
        print("Nenhum dataset carregado. Use a opção de importação primeiro.")
        return None
    for dataset in datasets:
        print(f"  [{dataset.id}] {dataset.identificador_exibicao()} - {dataset.df.shape[0]} linhas x {dataset.df.shape[1]} colunas")
    escolha = ler_inteiro(f"{mensagem}: ")
    dataset = sessao.obter_dataset(escolha)
    if dataset is None:
        print(f"Dataset {escolha} não encontrado.")
        return None
    return dataset


# ---------------------------------------------------------------------------
# Catálogo
# ---------------------------------------------------------------------------


def exibir_catalogo_geral(sessao: Sessao) -> None:
    """Exibe o catálogo completo (resumo e colunas) de todos os datasets carregados."""
    for dataset in sessao.listar_datasets():
        resumo = catalogo.montar_resumo_dataset(dataset.id, dataset.arquivo, dataset.aba, dataset.df)
        print(f"\n=== Dataset [{resumo['id']}] {resumo['arquivo']} :: {resumo['aba']} ===")
        print(f"Linhas: {resumo['linhas']}  Colunas: {resumo['colunas']}  Linhas duplicadas: {resumo['duplicadas']}")
        print(f"Possíveis colunas-chave: {resumo['colunas_chave']}")
        itens = catalogo.montar_catalogo_colunas(dataset.df)
        linhas_tabela = [
            {
                "#": item["posicao"],
                "Coluna": item["coluna"],
                "Tipo": item["tipo"],
                "Categoria": item["categoria"],
                "Preench.": item["preenchidos"],
                "Nulos": item["nulos"],
                "% Nulos": item["percentual_nulos"],
                "Únicos": item["unicos"],
                "Exemplos": truncar_texto(item["exemplos"], 40),
            }
            for item in itens
        ]
        _imprimir_tabela(linhas_tabela)


def exibir_previas_iniciais(sessao: Sessao) -> None:
    """Exibe uma prévia curta de cada dataset recém-carregado."""
    for dataset in sessao.listar_datasets():
        print(f"\nPrévia de {dataset.identificador_exibicao()}:")
        exibir_previa(dataset.df, linhas=5)


# ---------------------------------------------------------------------------
# Importação inicial de arquivos
# ---------------------------------------------------------------------------


def fluxo_importacao(sessao: Sessao) -> None:
    """Conduz o fluxo interativo de importação de um ou mais arquivos."""
    print("=== Importação de arquivos ===")
    algum_carregado = False
    while True:
        caminho_bruto = ler_texto("Informe o caminho do arquivo (XLSX, XLS, XLSM ou CSV): ")
        caminho = Path(caminho_bruto).expanduser()
        try:
            carregador.validar_arquivo(caminho)
            formato = carregador.identificar_formato(caminho)
            if formato in ("xlsx", "xlsm", "xls"):
                _importar_excel(sessao, caminho)
            else:
                _importar_csv(sessao, caminho)
            sessao.arquivos_processados.append(caminho.name)
            algum_carregado = True
        except ErroOperacao as exc:
            print(f"Erro: {exc}")
            logger.error("Falha ao importar '%s': %s", caminho_bruto, exc)
        except (KeyboardInterrupt, EOFError):
            raise
        except Exception as exc:  # noqa: BLE001 - qualquer falha de leitura não deve encerrar o programa
            print(f"Erro inesperado ao importar arquivo: {exc}")
            logger.exception("Erro inesperado ao importar '%s'", caminho_bruto)
        if not confirmar("Deseja importar outro arquivo?", padrao=False):
            break
    if not algum_carregado:
        print("Nenhum dataset foi carregado.")


def _importar_excel(sessao: Sessao, caminho: Path) -> None:
    """Lista as abas de um arquivo Excel e importa as abas escolhidas pelo usuário."""
    abas = carregador.listar_abas_excel(caminho)
    if not abas:
        print(f"Nenhuma aba encontrada em '{caminho.name}'.")
        return
    print(f"\nAbas encontradas em '{caminho.name}': {', '.join(abas)}")
    if confirmar("Importar todas as abas?", padrao=True):
        abas_escolhidas = abas
    else:
        for indice, aba in enumerate(abas, start=1):
            print(f"  [{indice}] {aba}")
        bruto = ler_texto("Informe os números das abas desejadas, separados por vírgula: ")
        indices = [int(p.strip()) for p in bruto.split(",") if p.strip().isdigit()]
        abas_escolhidas = [abas[i - 1] for i in indices if 1 <= i <= len(abas)]
        if not abas_escolhidas:
            print("Nenhuma aba válida selecionada.")
            return
    for aba in abas_escolhidas:
        try:
            df = carregador.carregar_planilha_excel(caminho, aba)
            dataset = sessao.adicionar_dataset(nome=f"{caminho.stem}_{aba}", arquivo=caminho.name, aba=aba, df=df)
            print(
                f"Dataset [{dataset.id}] carregado: '{caminho.name}' aba '{aba}' "
                f"({df.shape[0]} linhas x {df.shape[1]} colunas)"
            )
        except ErroOperacao as exc:
            print(f"Erro ao carregar a aba '{aba}': {exc}")


def _importar_csv(sessao: Sessao, caminho: Path) -> None:
    """Importa um arquivo CSV, detectando delimitador/codificação ou pedindo manualmente."""
    try:
        df, delimitador_usado, codificacao_usada = carregador.carregar_csv(caminho)
        print(f"CSV lido com delimitador '{delimitador_usado}' e codificação '{codificacao_usada}'.")
        if df.shape[1] == 1 and not confirmar(
            "Apenas uma coluna foi detectada; o delimitador pode estar incorreto. A leitura parece correta?",
            padrao=False,
        ):
            raise ErroOperacao("delimitador provavelmente incorreto")
    except ErroOperacao:
        print("Não foi possível detectar automaticamente o delimitador/codificação (ou o resultado não parece correto).")
        delimitador = ler_texto("Informe o delimitador manualmente (ex.: , ; | ou 'tab'): ", padrao=",")
        if delimitador.lower() == "tab":
            delimitador = "\t"
        codificacao = ler_texto("Informe a codificação manualmente (ex.: utf-8, latin-1, cp1252): ", padrao="utf-8")
        df, delimitador_usado, codificacao_usada = carregador.carregar_csv(caminho, delimitador, codificacao)
    dataset = sessao.adicionar_dataset(nome=caminho.stem, arquivo=caminho.name, aba=None, df=df)
    print(f"Dataset [{dataset.id}] carregado: '{caminho.name}' ({df.shape[0]} linhas x {df.shape[1]} colunas)")


# ---------------------------------------------------------------------------
# Fluxos padrão de confirmação (aplicar em dataset existente / criar novo / salvar resultado)
# ---------------------------------------------------------------------------


def aplicar_com_confirmacao(
    sessao: Sessao,
    dataset: Dataset,
    df_resultado: pd.DataFrame,
    operacao: str,
    parametros: str,
    avisos: Optional[list[str]] = None,
) -> bool:
    """Exibe a prévia do resultado e, mediante confirmação, aplica a alteração ao dataset."""
    print("\nPrévia do resultado (até 10 linhas):")
    exibir_previa(df_resultado)
    if avisos:
        print("\nAvisos:")
        for aviso in avisos:
            print(f"  - {aviso}")
    if not confirmar("Confirmar esta alteração?"):
        print("Operação cancelada. Nenhuma alteração foi aplicada.")
        return False

    linhas_antes, colunas_antes = dataset.df.shape
    dataset.salvar_estado_para_desfazer()
    dataset.df = df_resultado
    linhas_depois, colunas_depois = dataset.df.shape

    entrada = HistoricoEntry(
        timestamp=datetime.now(),
        dataset_id=dataset.id,
        dataset_nome=dataset.identificador_exibicao(),
        operacao=operacao,
        parametros=parametros,
        linhas_antes=linhas_antes,
        linhas_depois=linhas_depois,
        colunas_antes=colunas_antes,
        colunas_depois=colunas_depois,
        avisos="; ".join(avisos) if avisos else "",
    )
    sessao.registrar_historico(entrada)
    print("Alteração aplicada com sucesso.")
    return True


def criar_dataset_com_confirmacao(
    sessao: Sessao,
    nome_sugerido: str,
    df_resultado: pd.DataFrame,
    operacao: str,
    parametros: str,
    dataset_referencia: str = "-",
    avisos: Optional[list[str]] = None,
) -> bool:
    """Exibe a prévia e, mediante confirmação, cria um novo dataset derivado na sessão."""
    print("\nPrévia do novo dataset (até 10 linhas):")
    exibir_previa(df_resultado)
    if avisos:
        print("\nAvisos:")
        for aviso in avisos:
            print(f"  - {aviso}")
    if not confirmar("Confirmar a criação deste novo dataset?"):
        print("Operação cancelada.")
        return False

    novo_dataset = sessao.adicionar_dataset(nome=nome_sugerido, arquivo=nome_sugerido, aba=None, df=df_resultado, origem="derivado")
    entrada = HistoricoEntry(
        timestamp=datetime.now(),
        dataset_id=novo_dataset.id,
        dataset_nome=dataset_referencia,
        operacao=operacao,
        parametros=parametros,
        linhas_antes=0,
        linhas_depois=len(df_resultado),
        colunas_antes=0,
        colunas_depois=df_resultado.shape[1],
        avisos="; ".join(avisos) if avisos else "",
    )
    sessao.registrar_historico(entrada)
    print(f"Novo dataset [{novo_dataset.id}] '{nome_sugerido}' criado com sucesso.")
    return True


def salvar_resultado_com_confirmacao(
    sessao: Sessao,
    nome_sugerido: str,
    df_resultado: pd.DataFrame,
    operacao: str,
    parametros: str,
    dataset_referencia: str = "-",
    avisos: Optional[list[str]] = None,
) -> bool:
    """Exibe a prévia e, mediante confirmação, salva o resultado nomeado para exportação."""
    print("\nPrévia do resultado:")
    exibir_previa(df_resultado)
    if avisos:
        print("\nAvisos:")
        for aviso in avisos:
            print(f"  - {aviso}")
    if not confirmar("Deseja salvar este resultado para exportação no relatório final?", padrao=True):
        print("Resultado não foi salvo.")
        return False

    nome_final = sessao.adicionar_resultado(nome_sugerido, df_resultado)
    entrada = HistoricoEntry(
        timestamp=datetime.now(),
        dataset_id=0,
        dataset_nome=dataset_referencia,
        operacao=operacao,
        parametros=parametros,
        linhas_antes=0,
        linhas_depois=len(df_resultado),
        colunas_antes=0,
        colunas_depois=df_resultado.shape[1],
        avisos="; ".join(avisos) if avisos else "",
    )
    sessao.registrar_historico(entrada)
    print(f"Resultado salvo como '{nome_final}' para exportação.")
    return True


def _ler_pares_criterio(df: pd.DataFrame) -> list[tuple[str, str]]:
    """Lê pares (coluna, critério) combinados com E lógico, usados pelas funções *SES."""
    pares: list[tuple[str, str]] = []
    print("Informe os pares coluna/critério. Escolha 0 na coluna para terminar.")
    while True:
        coluna = selecionar_coluna(df, "Coluna do critério", permitir_cancelar=True)
        if coluna is None:
            break
        criterio = ler_texto(f"Critério para '{coluna}' (ex.: >100, <>0, abc*): ")
        pares.append((coluna, criterio))
        if not confirmar("Adicionar outro critério?", padrao=False):
            break
    return pares


# ---------------------------------------------------------------------------
# Handlers do menu principal
# ---------------------------------------------------------------------------


def handler_visualizar_catalogo(sessao: Sessao) -> None:
    """Opção 1: exibe o catálogo de datasets e colunas."""
    if not sessao.datasets:
        print("Nenhum dataset carregado.")
        return
    exibir_catalogo_geral(sessao)


def handler_visualizar_dados(sessao: Sessao) -> None:
    """Opção 2: exibe uma prévia dos dados e o resumo estatístico de um dataset."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    print(f"\n=== {dataset.identificador_exibicao()} ===")
    exibir_previa(dataset.df, linhas=15)
    print("\nResumo estatístico (colunas numéricas):")
    numericas = dataset.df.select_dtypes(include="number")
    if numericas.empty:
        print("(nenhuma coluna numérica)")
    else:
        with pd.option_context("display.width", 160):
            print(numericas.describe().to_string())


def handler_filtrar(sessao: Sessao) -> None:
    """Opção 3: filtra os dados combinando múltiplos critérios com E/OU."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    print(
        "Operadores disponíveis: igual, diferente, maior, maior_igual, menor, menor_igual, "
        "contem, nao_contem, comeca_com, termina_com, vazio, nao_vazio, em_lista, entre"
    )
    criterios: list[operacoes.Criterio] = []
    while True:
        coluna = selecionar_coluna(df, "Coluna do critério", permitir_cancelar=True)
        if coluna is None:
            break
        operador = ler_texto("Operador: ")
        valor: Any = None
        valor2: Any = None
        if operador not in ("vazio", "nao_vazio"):
            if operador == "em_lista":
                bruto = ler_texto("Valores (separados por vírgula): ")
                valor = [v.strip() for v in bruto.split(",")]
            elif operador == "entre":
                valor = ler_texto("Valor inicial: ")
                valor2 = ler_texto("Valor final: ")
            else:
                valor = ler_texto("Valor: ")
        criterios.append(operacoes.Criterio(coluna, operador, valor, valor2))
        if not confirmar("Adicionar outro critério?", padrao=False):
            break
    if not criterios:
        print("Nenhum critério informado.")
        return
    operador_logico = "E"
    if len(criterios) > 1:
        operador_logico = ler_texto("Combinar critérios com E ou OU: ", padrao="E").upper()
    try:
        resultado = operacoes.aplicar_filtros(df, criterios, operador_logico)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    parametros = f"criterios={[(c.coluna, c.operador, c.valor, c.valor2) for c in criterios]}, logico={operador_logico}"
    aplicar_com_confirmacao(sessao, dataset, resultado, "Filtrar dados", parametros)


def handler_ordenar(sessao: Sessao) -> None:
    """Opção 4: ordena os dados por uma ou várias colunas."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    colunas = selecionar_colunas_multiplas(df, "Colunas para ordenar (na ordem de prioridade)")
    if not colunas:
        print("Selecione ao menos uma coluna.")
        return
    ascendente = [confirmar(f"Ordem crescente para '{coluna}'?", padrao=True) for coluna in colunas]
    try:
        resultado = operacoes.ordenar_dados(df, colunas, ascendente)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, "Ordenar dados", f"colunas={colunas}, ascendente={ascendente}")


def handler_colunas(sessao: Sessao) -> None:
    """Opção 5: seleciona, reordena, renomeia ou remove colunas."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    print("  1. Selecionar colunas (manter apenas as escolhidas)")
    print("  2. Reordenar colunas")
    print("  3. Renomear colunas")
    print("  4. Remover colunas")
    escolha = ler_texto("Escolha: ")
    try:
        if escolha == "1":
            colunas = selecionar_colunas_multiplas(df, "Colunas a manter")
            if not colunas:
                print("Selecione ao menos uma coluna.")
                return
            resultado = operacoes.selecionar_colunas(df, colunas)
            parametros, nome_op = f"colunas={colunas}", "Selecionar colunas"
        elif escolha == "2":
            print("Ordem atual:")
            listar_colunas_numeradas(df)
            bruto = ler_texto("Informe a nova ordem (todos os números, separados por vírgula): ")
            indices = [int(p.strip()) for p in bruto.split(",") if p.strip().isdigit()]
            nova_ordem = [str(df.columns[i - 1]) for i in indices if 1 <= i <= len(df.columns)]
            resultado = operacoes.reordenar_colunas(df, nova_ordem)
            parametros, nome_op = f"ordem={nova_ordem}", "Reordenar colunas"
        elif escolha == "3":
            mapeamento: dict[str, str] = {}
            print("Informe as colunas a renomear. Escolha 0 para terminar.")
            while True:
                coluna = selecionar_coluna(df, "Coluna a renomear", permitir_cancelar=True)
                if coluna is None:
                    break
                novo_nome = ler_texto(f"Novo nome para '{coluna}': ")
                mapeamento[coluna] = novo_nome
                if not confirmar("Renomear outra coluna?", padrao=False):
                    break
            if not mapeamento:
                print("Nenhuma renomeação informada.")
                return
            resultado = operacoes.renomear_colunas(df, mapeamento)
            parametros, nome_op = f"mapeamento={mapeamento}", "Renomear colunas"
        elif escolha == "4":
            colunas = selecionar_colunas_multiplas(df, "Colunas a remover")
            if not colunas:
                print("Nenhuma coluna selecionada.")
                return
            if not confirmar(f"Confirma a remoção definitiva de {colunas}?", padrao=False):
                print("Operação cancelada.")
                return
            resultado = operacoes.remover_colunas(df, colunas)
            parametros, nome_op = f"colunas={colunas}", "Remover colunas"
        else:
            print("Opção inválida.")
            return
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, nome_op, parametros)


def handler_coluna_calculada(sessao: Sessao) -> None:
    """Opção 6: cria uma coluna calculada a partir de diversos tipos de cálculo."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    tipos = {
        "1": "Operação aritmética entre colunas (soma, subtração, multiplicação, divisão)",
        "2": "Percentual (parte / total * 100)",
        "3": "Condição lógica (SE)",
        "4": "Concatenação de colunas",
        "5": "Transformação de texto em uma coluna",
        "6": "Cálculo entre datas",
        "7": "Valor fixo",
        "8": "Expressão validada (envolvendo colunas numéricas)",
    }
    for chave, descricao in tipos.items():
        print(f"  {chave}. {descricao}")
    escolha = ler_texto("Escolha o tipo de coluna calculada: ")
    if escolha not in tipos:
        print("Opção inválida.")
        return
    nova_coluna = ler_texto("Nome da nova coluna: ")
    if nova_coluna in df.columns and not confirmar(f"A coluna '{nova_coluna}' já existe e será sobrescrita. Continuar?"):
        print("Operação cancelada.")
        return
    try:
        if escolha == "1":
            colunas = selecionar_colunas_multiplas(df, "Colunas numéricas (na ordem da operação)")
            operacao_aritmetica = ler_texto("Operação (soma/subtracao/multiplicacao/divisao): ")
            resultado = operacoes.criar_coluna_operacao_aritmetica(df, nova_coluna, colunas, operacao_aritmetica)
            parametros = f"colunas={colunas}, operacao={operacao_aritmetica}"
        elif escolha == "2":
            parte = selecionar_coluna(df, "Coluna da parte")
            total = selecionar_coluna(df, "Coluna do total")
            resultado = operacoes.criar_coluna_percentual(df, nova_coluna, parte, total)
            parametros = f"parte={parte}, total={total}"
        elif escolha == "3":
            condicao_texto = ler_texto("Expressão da condição (ex.: preco > 100): ")
            condicao = avaliar_condicao_segura(condicao_texto, df)
            valor_v = ler_texto("Valor se verdadeiro: ")
            valor_f = ler_texto("Valor se falso: ")
            resultado = operacoes.criar_coluna_condicional(df, nova_coluna, condicao, valor_v, valor_f)
            parametros = f"condicao={condicao_texto}, verdadeiro={valor_v}, falso={valor_f}"
        elif escolha == "4":
            colunas = selecionar_colunas_multiplas(df, "Colunas a concatenar")
            separador = ler_texto("Separador (Enter para nenhum): ", obrigatorio=False, padrao="")
            resultado = operacoes.criar_coluna_concatenacao(df, nova_coluna, colunas, separador)
            parametros = f"colunas={colunas}, separador='{separador}'"
        elif escolha == "5":
            coluna = selecionar_coluna(df, "Coluna de origem")
            transformacao = ler_texto("Transformação (maiusculas/minusculas/iniciais_maiusculas/remover_espacos): ")
            mapa_transformacoes = {
                "maiusculas": operacoes.texto_maiusculas,
                "minusculas": operacoes.texto_minusculas,
                "iniciais_maiusculas": operacoes.texto_iniciais_maiusculas,
                "remover_espacos": operacoes.texto_remover_espacos,
            }
            if transformacao not in mapa_transformacoes:
                print("Transformação inválida.")
                return
            resultado = df.copy()
            resultado[nova_coluna] = mapa_transformacoes[transformacao](df[coluna])
            parametros = f"coluna={coluna}, transformacao={transformacao}"
        elif escolha == "6":
            inicio = selecionar_coluna(df, "Coluna de data inicial")
            fim = selecionar_coluna(df, "Coluna de data final")
            unidade = ler_texto("Unidade (dias/horas/meses/anos): ", padrao="dias")
            resultado = operacoes.criar_coluna_diferenca_datas(df, nova_coluna, inicio, fim, unidade)
            parametros = f"inicio={inicio}, fim={fim}, unidade={unidade}"
        elif escolha == "7":
            valor = ler_texto("Valor fixo: ")
            resultado = operacoes.criar_coluna_valor_fixo(df, nova_coluna, valor)
            parametros = f"valor={valor}"
        else:
            expressao = ler_texto("Expressão (ex.: (preco - custo) / custo * 100): ")
            valores = avaliar_expressao_segura(expressao, df)
            resultado = operacoes.criar_coluna_expressao(df, nova_coluna, valores)
            parametros = f"expressao={expressao}"
    except (ErroOperacao, ErroExpressaoInsegura) as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, "Criar coluna calculada", parametros)


def handler_matematica(sessao: Sessao) -> None:
    """Opção 7: aplica uma operação matemática a uma coluna numérica."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    coluna = selecionar_coluna(df, "Coluna numérica de origem")
    opcoes = {
        "1": "somar_constante",
        "2": "subtrair_constante",
        "3": "multiplicar_constante",
        "4": "dividir_constante",
        "5": "potencia",
        "6": "raiz_quadrada",
        "7": "logaritmo",
        "8": "arredondar",
        "9": "absoluto",
    }
    for chave, descricao in opcoes.items():
        print(f"  {chave}. {descricao}")
    escolha = ler_texto("Escolha a operação: ")
    operacao_matematica = opcoes.get(escolha)
    if not operacao_matematica:
        print("Opção inválida.")
        return
    parametro = None
    if operacao_matematica in ("somar_constante", "subtrair_constante", "multiplicar_constante", "dividir_constante", "potencia", "arredondar"):
        parametro = ler_float("Informe o valor/parâmetro: ")
    nova_coluna = ler_texto("Nome da nova coluna: ")
    try:
        resultado = operacoes.aplicar_operacao_matematica(df, coluna, operacao_matematica, nova_coluna, parametro)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(
        sessao, dataset, resultado, "Operação matemática", f"coluna={coluna}, operacao={operacao_matematica}, parametro={parametro}"
    )


def handler_texto(sessao: Sessao) -> None:
    """Opção 8: aplica operações de texto a uma ou mais colunas."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    opcoes = {
        "1": "concatenar",
        "2": "esquerda",
        "3": "direita",
        "4": "localizar",
        "5": "substituir",
        "6": "remover_espacos",
        "7": "maiusculas",
        "8": "minusculas",
        "9": "iniciais_maiusculas",
        "10": "contar_caracteres",
        "11": "separar",
    }
    for chave, descricao in opcoes.items():
        print(f"  {chave}. {descricao}")
    escolha = ler_texto("Escolha a operação de texto: ")
    operacao_texto = opcoes.get(escolha)
    if not operacao_texto:
        print("Opção inválida.")
        return
    try:
        if operacao_texto == "concatenar":
            colunas = selecionar_colunas_multiplas(df, "Colunas a concatenar")
            separador = ler_texto("Separador (Enter para nenhum): ", obrigatorio=False, padrao="")
            nova_coluna = ler_texto("Nome da nova coluna: ")
            resultado = df.copy()
            resultado[nova_coluna] = operacoes.texto_concatenar(df, colunas, separador)
            parametros = f"colunas={colunas}, separador='{separador}'"
        elif operacao_texto == "separar":
            coluna = selecionar_coluna(df, "Coluna a separar")
            delimitador = ler_texto("Delimitador: ")
            expandido = operacoes.texto_separar(df[coluna], delimitador)
            expandido.columns = [f"{coluna}_{c}" for c in expandido.columns]
            resultado = pd.concat([df.reset_index(drop=True), expandido.reset_index(drop=True)], axis=1)
            parametros = f"coluna={coluna}, delimitador='{delimitador}'"
        else:
            coluna = selecionar_coluna(df, "Coluna de origem")
            nova_coluna = ler_texto("Nome da nova coluna: ")
            if operacao_texto == "esquerda":
                quantidade = ler_inteiro("Quantidade de caracteres: ", minimo=1)
                serie = operacoes.texto_esquerda(df[coluna], quantidade)
                parametros = f"coluna={coluna}, n={quantidade}"
            elif operacao_texto == "direita":
                quantidade = ler_inteiro("Quantidade de caracteres: ", minimo=1)
                serie = operacoes.texto_direita(df[coluna], quantidade)
                parametros = f"coluna={coluna}, n={quantidade}"
            elif operacao_texto == "localizar":
                subtexto = ler_texto("Texto a localizar: ")
                serie = operacoes.texto_localizar(df[coluna], subtexto)
                parametros = f"coluna={coluna}, subtexto='{subtexto}'"
            elif operacao_texto == "substituir":
                antigo = ler_texto("Texto a substituir: ")
                novo = ler_texto("Novo texto: ", obrigatorio=False, padrao="")
                serie = operacoes.texto_substituir(df[coluna], antigo, novo)
                parametros = f"coluna={coluna}, antigo='{antigo}', novo='{novo}'"
            elif operacao_texto == "remover_espacos":
                serie = operacoes.texto_remover_espacos(df[coluna])
                parametros = f"coluna={coluna}"
            elif operacao_texto == "maiusculas":
                serie = operacoes.texto_maiusculas(df[coluna])
                parametros = f"coluna={coluna}"
            elif operacao_texto == "minusculas":
                serie = operacoes.texto_minusculas(df[coluna])
                parametros = f"coluna={coluna}"
            elif operacao_texto == "iniciais_maiusculas":
                serie = operacoes.texto_iniciais_maiusculas(df[coluna])
                parametros = f"coluna={coluna}"
            else:
                serie = operacoes.texto_contar_caracteres(df[coluna])
                parametros = f"coluna={coluna}"
            resultado = df.copy()
            resultado[nova_coluna] = serie
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, "Operação de texto", parametros)


def handler_data(sessao: Sessao) -> None:
    """Opção 9: aplica operações de data a uma coluna."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    opcoes = {
        "1": "Converter coluna para data",
        "2": "Extrair componente (dia/mês/ano/trimestre/dia da semana)",
        "3": "Diferença entre duas colunas de data",
        "4": "Somar dias a uma coluna de data",
        "5": "Subtrair dias de uma coluna de data",
        "6": "Identificar datas inválidas",
        "7": "Criar coluna de período (mês/trimestre/ano) para agrupamento",
    }
    for chave, descricao in opcoes.items():
        print(f"  {chave}. {descricao}")
    escolha = ler_texto("Escolha a operação de data: ")
    try:
        if escolha == "1":
            coluna = selecionar_coluna(df, "Coluna a converter")
            resultado = df.copy()
            resultado[coluna] = operacoes.data_converter(df[coluna])
            parametros = f"coluna={coluna}"
        elif escolha == "2":
            coluna = selecionar_coluna(df, "Coluna de data")
            componente = ler_texto("Componente (dia/mes/ano/trimestre/dia_semana): ")
            nova_coluna = ler_texto("Nome da nova coluna: ")
            serie_data = operacoes.data_converter(df[coluna])
            resultado = df.copy()
            resultado[nova_coluna] = operacoes.data_extrair_componente(serie_data, componente)
            parametros = f"coluna={coluna}, componente={componente}"
        elif escolha == "3":
            inicio = selecionar_coluna(df, "Coluna de data inicial")
            fim = selecionar_coluna(df, "Coluna de data final")
            unidade = ler_texto("Unidade (dias/horas/meses/anos): ", padrao="dias")
            nova_coluna = ler_texto("Nome da nova coluna: ")
            resultado = operacoes.criar_coluna_diferenca_datas(df, nova_coluna, inicio, fim, unidade)
            parametros = f"inicio={inicio}, fim={fim}, unidade={unidade}"
        elif escolha in ("4", "5"):
            coluna = selecionar_coluna(df, "Coluna de data")
            dias = ler_inteiro("Quantidade de dias: ", minimo=0)
            if escolha == "5":
                dias = -dias
            nova_coluna = ler_texto("Nome da nova coluna: ")
            serie_data = operacoes.data_converter(df[coluna])
            resultado = df.copy()
            resultado[nova_coluna] = operacoes.data_somar_dias(serie_data, dias)
            parametros = f"coluna={coluna}, dias={dias}"
        elif escolha == "6":
            coluna = selecionar_coluna(df, "Coluna de data")
            nova_coluna = ler_texto("Nome da nova coluna: ")
            serie_data = operacoes.data_converter(df[coluna])
            resultado = df.copy()
            resultado[nova_coluna] = operacoes.data_identificar_invalidas(serie_data)
            parametros = f"coluna={coluna}"
        elif escolha == "7":
            coluna = selecionar_coluna(df, "Coluna de data")
            periodo = ler_texto("Período (mes/trimestre/ano): ")
            nova_coluna = ler_texto("Nome da nova coluna: ")
            resultado = operacoes.data_criar_coluna_periodo(df, coluna, periodo, nova_coluna)
            parametros = f"coluna={coluna}, periodo={periodo}"
        else:
            print("Opção inválida.")
            return
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, "Operação de data", parametros)


def handler_agrupar(sessao: Sessao) -> None:
    """Opção 10: agrupa e resume os dados, criando um novo dataset."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    colunas_grupo = selecionar_colunas_multiplas(df, "Colunas para agrupar (chave)")
    if not colunas_grupo:
        print("Selecione ao menos uma coluna de agrupamento.")
        return
    agregacoes: dict[str, str] = {}
    print("Selecione as colunas a resumir, uma por vez. Escolha 0 quando terminar.")
    while True:
        coluna = selecionar_coluna(df, "Coluna a resumir", permitir_cancelar=True)
        if coluna is None:
            break
        funcao = ler_texto("Função (soma/media/contagem/minimo/maximo/mediana/desvio_padrao/primeiro/ultimo): ")
        agregacoes[coluna] = funcao
        if not confirmar("Adicionar outra coluna a resumir?", padrao=False):
            break
    if not agregacoes:
        print("Nenhuma agregação informada.")
        return
    try:
        resultado = operacoes.agrupar_resumir(df, colunas_grupo, agregacoes)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    nome_sugerido = ler_texto("Nome para o novo dataset resumido: ", padrao=f"{dataset.nome}_resumo")
    criar_dataset_com_confirmacao(
        sessao, nome_sugerido, resultado, "Agrupar e resumir", f"grupo={colunas_grupo}, agregacoes={agregacoes}", dataset.identificador_exibicao()
    )


def handler_pivot(sessao: Sessao) -> None:
    """Opção 11: cria uma tabela dinâmica a partir de um dataset."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    linhas = selecionar_colunas_multiplas(df, "Colunas para LINHAS da tabela dinâmica")
    if not linhas:
        print("Selecione ao menos uma coluna para linhas.")
        return
    colunas_pivot = None
    if confirmar("Deseja definir colunas para o eixo COLUNAS da tabela?", padrao=False):
        colunas_pivot = selecionar_colunas_multiplas(df, "Colunas para COLUNAS da tabela dinâmica")
    valores = selecionar_colunas_multiplas(df, "Colunas de VALORES a agregar")
    if not valores:
        print("Selecione ao menos uma coluna de valores.")
        return
    funcao = ler_texto("Função de agregação (soma/media/contagem/minimo/maximo/mediana/desvio_padrao/primeiro/ultimo): ")
    preencher = None
    if confirmar("Deseja preencher células vazias com um valor?", padrao=False):
        preencher = ler_float("Valor para preenchimento: ")
    totais = confirmar("Incluir totais gerais?", padrao=False)
    try:
        resultado = operacoes.criar_tabela_dinamica(df, linhas, colunas_pivot, valores, funcao, preencher, totais)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    nome_resultado = ler_texto("Nome para esta tabela dinâmica: ", padrao=f"{dataset.nome}_pivot")
    salvar_resultado_com_confirmacao(
        sessao,
        nome_resultado,
        resultado,
        "Tabela dinâmica",
        f"linhas={linhas}, colunas={colunas_pivot}, valores={valores}, funcao={funcao}, preencher={preencher}, totais={totais}",
        dataset.identificador_exibicao(),
    )


def handler_procv(sessao: Sessao) -> None:
    """Opção 12: executa PROCV, enriquecendo o dataset principal com colunas do dataset de consulta."""
    print("Selecione o dataset PRINCIPAL (onde as colunas serão adicionadas):")
    principal = selecionar_dataset(sessao)
    if principal is None:
        return
    print("Selecione o dataset de CONSULTA (tabela de referência):")
    consulta = selecionar_dataset(sessao)
    if consulta is None:
        return
    chave_principal = selecionar_coluna(principal.df, "Coluna-chave no dataset principal")
    chave_consulta = selecionar_coluna(consulta.df, "Coluna-chave no dataset de consulta")
    colunas_retorno = selecionar_colunas_multiplas(consulta.df, "Colunas a retornar do dataset de consulta")
    if not colunas_retorno:
        print("Selecione ao menos uma coluna de retorno.")
        return
    sufixo = ler_texto("Sufixo para colunas em caso de conflito de nomes: ", padrao="_procv")
    try:
        resultado = operacoes.executar_procv(principal.df, consulta.df, chave_principal, chave_consulta, colunas_retorno, sufixo)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(
        sessao,
        principal,
        resultado.df,
        "PROCV",
        f"consulta={consulta.identificador_exibicao()}, chave_principal={chave_principal}, chave_consulta={chave_consulta}, retorno={colunas_retorno}",
        avisos=resultado.avisos,
    )


def handler_proch(sessao: Sessao) -> None:
    """Opção 13: executa PROCH (busca horizontal)."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    if len(df) == 0:
        print("O dataset está vazio.")
        return
    print(f"O dataset possui {len(df)} linhas (índices de 0 a {len(df) - 1}).")
    linha_chave = ler_inteiro("Linha (índice) que contém os valores de busca: ", minimo=0, maximo=len(df) - 1)
    valor_procurado = ler_texto("Valor procurado: ")
    linha_retorno = ler_inteiro("Linha (índice) da qual retornar o valor: ", minimo=0, maximo=len(df) - 1)
    try:
        resultado = operacoes.executar_proch(df, linha_chave, valor_procurado, linha_retorno)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    nome_resultado = ler_texto("Nome para salvar este resultado: ", padrao=f"{dataset.nome}_proch")
    salvar_resultado_com_confirmacao(
        sessao,
        nome_resultado,
        resultado.df,
        "PROCH",
        f"linha_chave={linha_chave}, valor_procurado={valor_procurado}, linha_retorno={linha_retorno}",
        dataset.identificador_exibicao(),
        avisos=resultado.avisos,
    )


def handler_procx(sessao: Sessao) -> None:
    """Opção 14: executa PROCX, enriquecendo o dataset principal com colunas do dataset de consulta."""
    print("Selecione o dataset PRINCIPAL:")
    principal = selecionar_dataset(sessao)
    if principal is None:
        return
    print("Selecione o dataset de CONSULTA:")
    consulta = selecionar_dataset(sessao)
    if consulta is None:
        return
    coluna_busca_principal = selecionar_coluna(principal.df, "Coluna de busca no dataset principal")
    coluna_busca_consulta = selecionar_coluna(consulta.df, "Coluna de busca no dataset de consulta")
    colunas_retorno = selecionar_colunas_multiplas(consulta.df, "Colunas a retornar")
    if not colunas_retorno:
        print("Selecione ao menos uma coluna de retorno.")
        return
    modo = ler_texto("Modo de correspondência (exata/aproximada): ", padrao="exata")
    ocorrencia = "primeira"
    valor_nao_encontrado = None
    if modo == "exata":
        ocorrencia = ler_texto("Em caso de múltiplas ocorrências, usar (primeira/ultima): ", padrao="primeira")
        if confirmar("Deseja definir um valor para quando não houver correspondência?", padrao=False):
            valor_nao_encontrado = ler_texto("Valor para 'não encontrado': ")
    try:
        resultado = operacoes.executar_procx(
            principal.df, consulta.df, coluna_busca_principal, coluna_busca_consulta, colunas_retorno, modo, ocorrencia, valor_nao_encontrado
        )
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(
        sessao,
        principal,
        resultado.df,
        "PROCX",
        (
            f"consulta={consulta.identificador_exibicao()}, busca_principal={coluna_busca_principal}, "
            f"busca_consulta={coluna_busca_consulta}, retorno={colunas_retorno}, modo={modo}, ocorrencia={ocorrencia}"
        ),
        avisos=resultado.avisos,
    )


def handler_somase(sessao: Sessao) -> None:
    """Opção 15: calcula SOMASE."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    coluna_soma = selecionar_coluna(df, "Coluna a somar")
    coluna_criterio = selecionar_coluna(df, "Coluna do critério")
    criterio = ler_texto("Critério (ex.: >100, <>0, abc*, texto): ")
    try:
        valor = operacoes.somase(df, coluna_soma, coluna_criterio, criterio)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    print(f"\nResultado SOMASE: {valor}")
    resultado_df = pd.DataFrame([{"Função": "SOMASE", "Coluna soma": coluna_soma, "Critério": f"{coluna_criterio} {criterio}", "Resultado": valor}])
    nome_resultado = ler_texto("Nome para salvar este resultado: ", padrao=f"{dataset.nome}_somase")
    salvar_resultado_com_confirmacao(
        sessao, nome_resultado, resultado_df, "SOMASE", f"coluna_soma={coluna_soma}, coluna_criterio={coluna_criterio}, criterio={criterio}",
        dataset.identificador_exibicao(),
    )


def handler_somases(sessao: Sessao) -> None:
    """Opção 16: calcula SOMASES."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    coluna_soma = selecionar_coluna(df, "Coluna a somar")
    pares = _ler_pares_criterio(df)
    if not pares:
        print("Nenhum critério informado.")
        return
    try:
        valor = operacoes.somases(df, coluna_soma, pares)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    print(f"\nResultado SOMASES: {valor}")
    resultado_df = pd.DataFrame([{"Função": "SOMASES", "Coluna soma": coluna_soma, "Critérios": str(pares), "Resultado": valor}])
    nome_resultado = ler_texto("Nome para salvar este resultado: ", padrao=f"{dataset.nome}_somases")
    salvar_resultado_com_confirmacao(
        sessao, nome_resultado, resultado_df, "SOMASES", f"coluna_soma={coluna_soma}, criterios={pares}", dataset.identificador_exibicao()
    )


def handler_contse(sessao: Sessao) -> None:
    """Opção 17: calcula CONT.SE."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    coluna_criterio = selecionar_coluna(df, "Coluna do critério")
    criterio = ler_texto("Critério (ex.: >100, <>0, abc*, texto): ")
    try:
        valor = operacoes.cont_se(df, coluna_criterio, criterio)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    print(f"\nResultado CONT.SE: {valor}")
    resultado_df = pd.DataFrame([{"Função": "CONT.SE", "Critério": f"{coluna_criterio} {criterio}", "Resultado": valor}])
    nome_resultado = ler_texto("Nome para salvar este resultado: ", padrao=f"{dataset.nome}_contse")
    salvar_resultado_com_confirmacao(
        sessao, nome_resultado, resultado_df, "CONT.SE", f"coluna_criterio={coluna_criterio}, criterio={criterio}", dataset.identificador_exibicao()
    )


def handler_contses(sessao: Sessao) -> None:
    """Opção 18: calcula CONT.SES."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    pares = _ler_pares_criterio(df)
    if not pares:
        print("Nenhum critério informado.")
        return
    try:
        valor = operacoes.cont_ses(df, pares)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    print(f"\nResultado CONT.SES: {valor}")
    resultado_df = pd.DataFrame([{"Função": "CONT.SES", "Critérios": str(pares), "Resultado": valor}])
    nome_resultado = ler_texto("Nome para salvar este resultado: ", padrao=f"{dataset.nome}_contses")
    salvar_resultado_com_confirmacao(sessao, nome_resultado, resultado_df, "CONT.SES", f"criterios={pares}", dataset.identificador_exibicao())


def handler_mediase(sessao: Sessao) -> None:
    """Opção 19: calcula MÉDIASE."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    coluna_media = selecionar_coluna(df, "Coluna a calcular a média")
    coluna_criterio = selecionar_coluna(df, "Coluna do critério")
    criterio = ler_texto("Critério (ex.: >100, <>0, abc*, texto): ")
    try:
        valor = operacoes.mediase(df, coluna_media, coluna_criterio, criterio)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    print(f"\nResultado MÉDIASE: {valor}")
    resultado_df = pd.DataFrame([{"Função": "MÉDIASE", "Coluna média": coluna_media, "Critério": f"{coluna_criterio} {criterio}", "Resultado": valor}])
    nome_resultado = ler_texto("Nome para salvar este resultado: ", padrao=f"{dataset.nome}_mediase")
    salvar_resultado_com_confirmacao(
        sessao, nome_resultado, resultado_df, "MÉDIASE", f"coluna_media={coluna_media}, coluna_criterio={coluna_criterio}, criterio={criterio}",
        dataset.identificador_exibicao(),
    )


def handler_mediases(sessao: Sessao) -> None:
    """Opção 20: calcula MÉDIASES."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    coluna_media = selecionar_coluna(df, "Coluna a calcular a média")
    pares = _ler_pares_criterio(df)
    if not pares:
        print("Nenhum critério informado.")
        return
    try:
        valor = operacoes.mediases(df, coluna_media, pares)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    print(f"\nResultado MÉDIASES: {valor}")
    resultado_df = pd.DataFrame([{"Função": "MÉDIASES", "Coluna média": coluna_media, "Critérios": str(pares), "Resultado": valor}])
    nome_resultado = ler_texto("Nome para salvar este resultado: ", padrao=f"{dataset.nome}_mediases")
    salvar_resultado_com_confirmacao(
        sessao, nome_resultado, resultado_df, "MÉDIASES", f"coluna_media={coluna_media}, criterios={pares}", dataset.identificador_exibicao()
    )


def handler_se(sessao: Sessao) -> None:
    """Opção 21: cria coluna equivalente à função SE."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    condicao_texto = ler_texto("Condição (ex.: idade >= 18): ")
    try:
        condicao = avaliar_condicao_segura(condicao_texto, df)
    except ErroExpressaoInsegura as exc:
        print(f"Erro: {exc}")
        return
    valor_v = ler_texto("Valor se verdadeiro: ")
    valor_f = ler_texto("Valor se falso: ")
    nova_coluna = ler_texto("Nome da nova coluna: ")
    resultado = operacoes.funcao_se(df, nova_coluna, condicao, valor_v, valor_f)
    aplicar_com_confirmacao(sessao, dataset, resultado, "SE", f"condicao={condicao_texto}, verdadeiro={valor_v}, falso={valor_f}")


def handler_se_aninhado(sessao: Sessao) -> None:
    """Opção 22: cria coluna equivalente à função SE aninhado."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    condicoes_texto: list[str] = []
    condicoes: list[pd.Series] = []
    valores: list[Any] = []
    print("Informe as condições em ordem de prioridade. Deixe em branco para terminar.")
    while True:
        condicao_texto = ler_texto("Condição (Enter para terminar): ", obrigatorio=False, padrao="")
        if not condicao_texto:
            break
        try:
            condicao = avaliar_condicao_segura(condicao_texto, df)
        except ErroExpressaoInsegura as exc:
            print(f"Erro: {exc}")
            continue
        valor = ler_texto(f"Valor se '{condicao_texto}' for verdadeira: ")
        condicoes_texto.append(condicao_texto)
        condicoes.append(condicao)
        valores.append(valor)
    if not condicoes:
        print("Nenhuma condição informada.")
        return
    valor_padrao = ler_texto("Valor padrão (se nenhuma condição for verdadeira): ")
    nova_coluna = ler_texto("Nome da nova coluna: ")
    try:
        resultado = operacoes.funcao_se_aninhado(df, nova_coluna, condicoes, valores, valor_padrao)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, "SE aninhado", f"condicoes={condicoes_texto}, valores={valores}, padrao={valor_padrao}")


def handler_nulos(sessao: Sessao) -> None:
    """Opção 23: remove ou preenche valores nulos."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    print("  1. Remover linhas com valores nulos")
    print("  2. Preencher valores nulos")
    escolha = ler_texto("Escolha: ")
    try:
        if escolha == "1":
            todas = confirmar("Considerar todas as colunas? (Não, para escolher colunas específicas)", padrao=True)
            colunas = None if todas else selecionar_colunas_multiplas(df, "Colunas a considerar")
            resultado = operacoes.remover_linhas_nulas(df, colunas)
            parametros, nome_op = f"colunas={colunas or 'todas'}", "Remover linhas nulas"
        elif escolha == "2":
            colunas = selecionar_colunas_multiplas(df, "Colunas a preencher")
            if not colunas:
                print("Selecione ao menos uma coluna.")
                return
            metodo = ler_texto("Método (valor/media/mediana/moda/zero/texto): ")
            valor = ler_texto("Valor a utilizar: ") if metodo in ("valor", "texto") else None
            resultado = operacoes.preencher_nulos(df, colunas, metodo, valor)
            parametros, nome_op = f"colunas={colunas}, metodo={metodo}, valor={valor}", "Preencher valores nulos"
        else:
            print("Opção inválida.")
            return
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, nome_op, parametros)


def handler_duplicados(sessao: Sessao) -> None:
    """Opção 24: remove valores duplicados."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    df = dataset.df
    todas = confirmar("Considerar todas as colunas para identificar duplicidade?", padrao=True)
    colunas = None if todas else selecionar_colunas_multiplas(df, "Colunas a considerar")
    manter = ler_texto("Manter qual ocorrência (primeira/ultima): ", padrao="primeira")
    try:
        resultado = operacoes.remover_duplicados(df, colunas, manter)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    aplicar_com_confirmacao(sessao, dataset, resultado, "Remover duplicados", f"colunas={colunas or 'todas'}, manter={manter}")


def handler_concatenar(sessao: Sessao) -> None:
    """Opção 25: concatena (empilha) vários datasets, criando um novo."""
    if len(sessao.datasets) < 2:
        print("São necessários ao menos dois datasets carregados.")
        return
    print("Selecione os datasets a concatenar (informe os IDs separados por vírgula):")
    for dataset in sessao.listar_datasets():
        print(f"  [{dataset.id}] {dataset.identificador_exibicao()}")
    bruto = ler_texto("IDs: ")
    ids = [int(p.strip()) for p in bruto.split(",") if p.strip().isdigit()]
    datasets = [sessao.obter_dataset(i) for i in ids]
    if len(datasets) < 2 or any(d is None for d in datasets):
        print("Selecione ao menos dois datasets válidos.")
        return
    try:
        resultado = operacoes.concatenar_datasets([d.df for d in datasets])  # type: ignore[union-attr]
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    nome = ler_texto("Nome para o novo dataset concatenado: ", padrao="concatenado")
    criar_dataset_com_confirmacao(
        sessao, nome, resultado, "Concatenar datasets", f"datasets={ids}",
        ", ".join(d.identificador_exibicao() for d in datasets),  # type: ignore[union-attr]
    )


def handler_mesclar(sessao: Sessao) -> None:
    """Opção 26: mescla dois datasets por uma chave, criando um novo dataset."""
    print("Selecione o dataset ESQUERDO:")
    esquerdo = selecionar_dataset(sessao)
    if esquerdo is None:
        return
    print("Selecione o dataset DIREITO:")
    direito = selecionar_dataset(sessao)
    if direito is None:
        return
    chave_esquerda = selecionar_coluna(esquerdo.df, "Coluna-chave do dataset esquerdo")
    chave_direita = selecionar_coluna(direito.df, "Coluna-chave do dataset direito")
    tipo = ler_texto("Tipo de junção (inner/left/right/outer): ", padrao="inner")
    try:
        resultado = operacoes.mesclar_datasets(esquerdo.df, direito.df, chave_esquerda, chave_direita, tipo)
    except ErroOperacao as exc:
        print(f"Erro: {exc}")
        return
    nome = ler_texto("Nome para o novo dataset mesclado: ", padrao="mesclado")
    criar_dataset_com_confirmacao(
        sessao,
        nome,
        resultado,
        "Mesclar datasets",
        f"esquerda={esquerdo.identificador_exibicao()}, direita={direito.identificador_exibicao()}, chave_esq={chave_esquerda}, chave_dir={chave_direita}, tipo={tipo}",
        f"{esquerdo.identificador_exibicao()} + {direito.identificador_exibicao()}",
    )


def handler_qualidade(sessao: Sessao) -> None:
    """Opção 27: valida a qualidade dos dados de um dataset."""
    dataset = selecionar_dataset(sessao)
    if dataset is None:
        return
    qualidade = operacoes.validar_qualidade_dados(dataset.df)
    print(f"\n=== Qualidade dos dados: {dataset.identificador_exibicao()} ===")
    print(f"Linhas: {qualidade['total_linhas']}  Colunas: {qualidade['total_colunas']}")
    print(f"Total de nulos: {qualidade['total_nulos']} ({qualidade['percentual_nulos']}%)")
    print(f"Linhas duplicadas: {qualidade['linhas_duplicadas']}")
    print(f"Colunas totalmente nulas: {', '.join(qualidade['colunas_totalmente_nulas']) or 'nenhuma'}")
    print(f"Colunas constantes (valor único): {', '.join(qualidade['colunas_constantes']) or 'nenhuma'}")
    relatorio_df = catalogo.montar_relatorio_qualidade(dataset.df)
    if confirmar("Deseja salvar este relatório de qualidade para exportação?", padrao=True):
        nome_final = sessao.adicionar_resultado(f"{dataset.nome}_qualidade", relatorio_df)
        print(f"Relatório salvo como '{nome_final}' para exportação.")


def handler_desfazer(sessao: Sessao) -> None:
    """Opção 28: desfaz a última operação confirmada em um dataset."""
    dataset = selecionar_dataset(sessao, "Escolha o dataset para desfazer a última operação")
    if dataset is None:
        return
    if not dataset.pode_desfazer():
        print("Não há operações para desfazer neste dataset.")
        return
    linhas_antes, colunas_antes = dataset.df.shape
    dataset.desfazer()
    linhas_depois, colunas_depois = dataset.df.shape
    entrada = HistoricoEntry(
        timestamp=datetime.now(),
        dataset_id=dataset.id,
        dataset_nome=dataset.identificador_exibicao(),
        operacao="Desfazer última operação",
        parametros="-",
        linhas_antes=linhas_antes,
        linhas_depois=linhas_depois,
        colunas_antes=colunas_antes,
        colunas_depois=colunas_depois,
    )
    sessao.registrar_historico(entrada)
    print("Última operação desfeita com sucesso.")
    exibir_previa(dataset.df)


def handler_historico(sessao: Sessao) -> None:
    """Opção 29: exibe o histórico de operações confirmadas."""
    if not sessao.historico:
        print("Nenhuma operação registrada ainda.")
        return
    _imprimir_tabela([entrada.como_linha() for entrada in sessao.historico])


def handler_salvar(sessao: Sessao) -> None:
    """Opção 30: gera o relatório final em Excel."""
    if not sessao.datasets:
        print("Nenhum dataset carregado; não há o que exportar.")
        return
    caminho = Path(ler_texto("Nome do arquivo de saída: ", padrao="relatorio.xlsx"))
    if not caminho.suffix:
        caminho = caminho.with_suffix(".xlsx")
    if caminho.exists():
        if not confirmar(f"O arquivo '{caminho}' já existe. Deseja sobrescrevê-lo?", padrao=False):
            caminho = exportador.gerar_nome_com_timestamp(caminho)
            print(f"Um novo nome será utilizado: {caminho}")
    try:
        caminho_final = exportador.gerar_relatorio(sessao, caminho)
    except Exception as exc:  # noqa: BLE001 - qualquer falha na exportação não deve encerrar o programa
        print(f"Erro ao gerar o relatório: {exc}")
        logger.exception("Falha ao gerar relatório")
        return
    print(f"Relatório salvo em: {caminho_final.resolve()}")
