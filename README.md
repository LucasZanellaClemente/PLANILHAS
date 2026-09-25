# Ferramenta de Análise de Planilhas em Python

Aplicação de linha de comando, em Python 3.11+, que importa planilhas (XLSX,
XLS, XLSM ou CSV), monta um catálogo detalhado de datasets e colunas, e
disponibiliza um menu interativo com operações equivalentes às principais
funções do Excel (PROCV, PROCX, SOMASE, tabela dinâmica, filtros, colunas
calculadas etc.). Ao final, exporta os dados, as análises e o histórico de
operações para `relatorio.xlsx`.

## Estrutura do projeto

```
excel_toolkit/
├── main.py           # Ponto de entrada e laço do menu principal
├── interface.py       # Leitura de entrada do usuário e handlers de cada opção do menu
├── carregador.py       # Importação e validação de arquivos (XLSX/XLS/XLSM/CSV)
├── catalogo.py         # Construção do catálogo de datasets e colunas
├── operacoes.py         # Regras de negócio: filtros, PROCV/PROCH/PROCX, SOMASE etc.
├── exportador.py        # Geração do relatorio.xlsx formatado (XlsxWriter)
├── estado.py             # Modelos de dados: Dataset, Sessao, HistoricoEntry
├── utils.py                # Utilitários: avaliador seguro de expressões, sanitização, curingas
├── requirements.txt
└── README.md
```

**Separação de responsabilidades:** `interface.py` cuida apenas da conversa
com o usuário (perguntas, prévias, confirmações); `operacoes.py` contém
funções puras que recebem DataFrames e parâmetros já validados e devolvem
novos DataFrames, sem nunca alterar os dados originais; `exportador.py`
conhece apenas como transformar o estado da sessão em um arquivo Excel
formatado; `estado.py` define os modelos de dados compartilhados.

## Instalação

Requer Python 3.11 ou superior.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Em Linux/macOS, ative o ambiente virtual com `source .venv/bin/activate`.

## Como executar

```bash
python main.py
```

## Exemplo de uso

```
============================================================
Ferramenta de Análise de Planilhas em Python
Formatos suportados: XLSX, XLS, XLSM, CSV
============================================================
=== Importação de arquivos ===
Informe o caminho do arquivo (XLSX, XLS, XLSM ou CSV): vendas.csv
CSV lido com delimitador ',' e codificação 'utf-8'.
Dataset [1] carregado: 'vendas.csv' (8 linhas x 7 colunas)
Deseja importar outro arquivo? [s/N]: n

=== Catálogo dos datasets carregados ===

=== Dataset [1] vendas.csv :: - ===
Linhas: 8  Colunas: 7  Linhas duplicadas: 0
Possíveis colunas-chave: id
  #  Coluna        Tipo       Categoria  Preench.  Nulos  % Nulos  Únicos  Exemplos
  1  id            numérico   numérica          8      0     0.00       8  1, 2, 3
  2  produto       texto      textual           7      1    12.50       4  Caneta, Caderno
  ...

============================================================
MENU PRINCIPAL - Ferramenta de Análise de Planilhas
============================================================
  0. Sair
  1. Visualizar catálogo das colunas
  ...
  30. Salvar projeto ou relatório

Escolha uma opção: 3
[...menu de filtro...]

Escolha uma opção: 30
Nome do arquivo de saída: relatorio.xlsx
Relatório salvo em: C:\...\relatorio.xlsx
```

O `relatorio.xlsx` gerado contém, entre outras, as abas `Resumo`,
`Catalogo_Colunas`, `Qualidade_Dados`, `Historico`, uma aba por dataset
original (`<nome>_orig`), uma aba por dataset alterado (`<nome>_alt`) e uma
aba por resultado nomeado (tabelas dinâmicas, PROCH, SOMASE etc.).

## Arquitetura (resumo)

- **`carregador.py`** identifica o formato do arquivo pela extensão, detecta
  delimitador (via `csv.Sniffer`, com contagem de ocorrências como
  alternativa) e codificação (testando `utf-8-sig`, `utf-8`, `cp1252` e
  `latin-1`) de arquivos CSV, e lista/lê abas de arquivos Excel com
  `pandas` + `openpyxl` (ou `xlrd` para `.xls`).
- **`catalogo.py`** classifica cada coluna (numérica, texto, data, booleana),
  sugere colunas-chave (sem nulos e 100% únicas) e monta os metadados
  exibidos no catálogo.
- **`operacoes.py`** é o núcleo de regras de negócio: cada operação Excel
  (PROCV, PROCH, PROCX, SOMASE/SOMASES, CONT.SE/CONT.SES, MÉDIASE/MÉDIASES,
  SE/SE aninhado, filtros, ordenação, tabela dinâmica, tratamento de nulos e
  duplicados etc.) é uma função pura que recebe DataFrame(s) e parâmetros e
  devolve um novo DataFrame, levantando `ErroOperacao` em casos inválidos.
  Nenhuma função altera o DataFrame recebido — todas trabalham em cópias.
- **`utils.py`** implementa o avaliador seguro de expressões (usado nas
  colunas calculadas e nas funções SE/SE aninhado): a expressão do usuário é
  convertida em uma árvore `ast`, validada nó a nó (apenas operadores
  aritméticos/lógicos, nomes de colunas existentes e uma lista restrita de
  funções matemáticas são aceitos) e só então avaliada — nunca com `eval()`
  direto sobre a string bruta.
- **`interface.py`** conduz toda a conversa com o usuário: seleção de
  datasets/colunas por número, leitura validada de inteiros/floats/textos,
  fluxo de importação inicial, e um handler por opção do menu. Após cada
  operação que altera um dataset, `aplicar_com_confirmacao` mostra uma
  prévia, pede confirmação, salva o estado anterior na pilha de desfazer e
  só então registra a operação no histórico. Operações que geram um
  resultado avulso (tabela dinâmica, PROCH, SOMASE etc.) usam
  `salvar_resultado_com_confirmacao`, que guarda o resultado nomeado para
  exportação. Operações que combinam datasets (concatenar, mesclar,
  agrupar) usam `criar_dataset_com_confirmacao`, que registra um novo
  dataset derivado, reutilizável nas demais operações do menu.
- **`estado.py`** define `Dataset` (dados atuais + cópia original imutável +
  pilha de desfazer com limite configurável), `Sessao` (registro central de
  datasets, resultados nomeados e histórico) e `HistoricoEntry`.
- **`exportador.py`** usa `pandas.ExcelWriter` com engine `xlsxwriter` para
  gerar o relatório final: cabeçalhos destacados, primeira linha congelada,
  autofiltro (ou tabela nativa do Excel quando há dados), largura de coluna
  ajustada ao conteúdo, formatos numérico/data/percentual, nomes de aba
  únicos e válidos (até 31 caracteres) e sanitização de valores que
  começam com `=`, `+`, `-` ou `@` (prefixados com apóstrofo) para evitar
  injeção de fórmulas.

### Decisões de design relevantes

- **PROCV/PROCX** enriquecem diretamente o dataset "principal" com as
  colunas retornadas (como uma fórmula do Excel arrastada por uma coluna),
  respeitando o fluxo padrão de confirmação/desfazer.
- **Tabela dinâmica, PROCH, SOMASE/SOMASES/CONT.SE/CONT.SES/MÉDIASE/MÉDIASES**
  geram um *resultado nomeado* (não um dataset editável), pois representam
  uma consulta/agregação pontual — exatamente como pedido na especificação
  do relatório final ("uma aba para cada tabela dinâmica" e "abas com os
  resultados de buscas e agregações"). Eles são salvos como abas próprias no
  `relatorio.xlsx`, com o nome escolhido pelo usuário.
- **Concatenar e Mesclar datasets** e **Agrupar e resumir** criam um novo
  *dataset* completo (não apenas um resultado), pois seu formato é idêntico
  ao de qualquer outro dataset e faz sentido continuar filtrando/ordenando/
  transformando esse resultado pelo próprio menu.
- **Desfazer** mantém uma pilha de até `limite_desfazer` (padrão 20) estados
  anteriores por dataset; cada `desfazer` remove o topo da pilha.

## Principais limitações

- **PROCH** (busca horizontal) é implementado varrendo as colunas de duas
  linhas específicas do DataFrame (uma para os valores de busca, outra para
  o retorno), já que o pandas é otimizado para operações por coluna, não por
  linha — a busca é O(n) por natureza, diferente de um lookup indexado.
- **PROCX** não distingue estruturalmente "busca à esquerda" de "busca à
  direita": como as colunas são referenciadas pelo nome (não pela posição
  relativa a um intervalo), qualquer coluna do dataset de consulta pode ser
  retornada, independentemente de sua posição — essa flexibilidade já
  reproduz (e amplia) o comportamento do XLOOKUP nesse aspecto.
- Tabelas dinâmicas, PROCH e as funções `*SE`/`*SES` **não geram um dataset
  editável** pelo menu (por decisão de design — ver seção acima); para
  continuar transformando esse resultado, seria necessário reimportá-lo ou
  reconstruí-lo.
- A detecção de codificação de CSV testa um conjunto fixo de codificações
  comuns (`utf-8-sig`, `utf-8`, `cp1252`, `latin-1`); arquivos em
  codificações incomuns podem exigir informar a codificação manualmente.
- O avaliador seguro de expressões aceita apenas operadores aritméticos,
  comparações, `E`/`OU`/`NÃO` (via `and`/`or`/`not`, ou `&`/`|`/`~` com cada
  condição entre parênteses) e um conjunto fixo de
  funções (`abs`, `round`, `min`, `max`, `sqrt`, `log`, `log10`, `exp`) —
  expressões mais complexas do Excel não têm equivalente direto.
- A aplicação é de linha de comando (sem interface gráfica) e de sessão
  única: o estado (datasets, histórico) vive apenas em memória durante a
  execução e não é persistido entre execuções, exceto pelo `relatorio.xlsx`
  exportado.
- Arquivos muito grandes podem consumir bastante memória, pois todas as
  cópias usadas para permitir "desfazer" ficam em memória (limitadas por
  dataset ao parâmetro `limite_desfazer`).

## Checklist de testes manuais

- [ ] Importar um CSV com delimitador `,` e outro com `;`, verificando a
      detecção automática de delimitador/codificação.
- [ ] Importar um CSV forçando delimitador/codificação manualmente (opção
      de fallback) e confirmar que a leitura fica correta.
- [ ] Importar um XLSX com múltiplas abas, escolhendo "importar todas" em
      uma execução e "escolher abas específicas" em outra.
- [ ] Verificar o catálogo (opção 1): tipos de coluna, contagem de nulos,
      percentuais, valores únicos e colunas-chave sugeridas.
- [ ] Aplicar um filtro com um único critério e outro combinando dois
      critérios com `E` e depois com `OU`.
- [ ] Ordenar por uma coluna e por múltiplas colunas com direções
      diferentes.
- [ ] Selecionar, reordenar, renomear e remover colunas (confirmando a
      remoção).
- [ ] Criar uma coluna calculada de cada tipo (aritmética, percentual,
      condição lógica, concatenação, transformação de texto, diferença de
      datas, valor fixo, expressão validada) e testar uma expressão
      inválida/insegura para confirmar que é rejeitada.
- [ ] Aplicar uma operação matemática (ex.: arredondar) e uma de texto
      (ex.: iniciais maiúsculas) e uma de data (ex.: extrair mês).
- [ ] Agrupar e resumir por uma coluna, e criar uma tabela dinâmica com
      linhas, colunas e valores.
- [ ] Executar PROCV e PROCX entre dois datasets, incluindo um caso com
      chaves duplicadas ou sem correspondência, e verificar os avisos.
- [ ] Executar PROCH informando uma linha de busca e uma linha de retorno.
- [ ] Executar SOMASE, SOMASES, CONT.SE, CONT.SES, MÉDIASE e MÉDIASES com
      critérios numéricos, de texto e com curinga (`*`).
- [ ] Criar uma coluna com SE e outra com SE aninhado (múltiplas condições).
- [ ] Remover linhas nulas, preencher nulos (por valor, média, mediana,
      moda, zero e texto) e remover duplicados (mantendo primeira/última
      ocorrência).
- [ ] Concatenar dois datasets e mesclar dois datasets com diferentes tipos
      de junção (inner/left/right/outer).
- [ ] Validar a qualidade dos dados de um dataset e salvar o relatório de
      qualidade.
- [ ] Desfazer a última operação de um dataset e confirmar que os dados
      voltam ao estado anterior.
- [ ] Visualizar o histórico de operações após várias alterações.
- [ ] Salvar o relatório, confirmar a sobrescrita de um arquivo existente e,
      em outra execução, recusar a sobrescrita e verificar que um novo nome
      com data/hora é gerado.
- [ ] Abrir o `relatorio.xlsx` gerado e confirmar: abas obrigatórias
      presentes, cabeçalhos destacados, primeira linha congelada,
      autofiltro/tabela, larguras de coluna ajustadas e que nenhuma célula
      contém uma fórmula não intencional (testar importando um CSV com um
      valor começando por `=`).
