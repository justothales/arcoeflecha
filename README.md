# Gerador de Combates para Competições de Arco e Flecha (Robin Round)

Este projeto é uma ferramenta web desenvolvida para automatizar a criação de chaves de combate para torneios de arco e flecha que utilizam o formato *Robin Round Individual*. A aplicação foi migrada de um script Jupyter Notebook para uma interface web interativa e amigável usando a biblioteca Streamlit.

## Entradas (Inputs)

A aplicação requer três tipos de entrada fornecidos pelo usuário através da interface web:

1.  **Informações da Prova:**
    *   **Nome da Etapa:** Campo de texto para identificar o nome do torneio (ex: "7º Outdoor FPAF - 4º Robin Round Individual").
    *   **Local da Prova:** Campo de texto para identificar onde a prova está sendo realizada (ex: "Mairiporã").

2.  **Planilha de Resultados da Prova (.xlsx):**
    *   Um arquivo Excel contendo a classificação inicial dos atletas após uma rodada qualificatória.
    *   **Colunas essenciais:** `total`, `DivisionClass` (categoria do atleta), `Rank` (classificação na categoria), `NomeCompleto` e `local_prova` (clube do atleta).

3.  **Planilha de Distribuição de Grupos (.xlsx):**
    *   Um arquivo Excel chamado `DistGrupos.xlsx` que funciona como uma matriz de mapeamento. Ele determina para qual grupo um atleta será alocado com base no número total de competidores na sua categoria e no seu `Rank` individual.

## Processamento

O coração do script realiza uma série de manipulações nos dados para criar os grupos e os combates de forma justa e automática:

1.  **Filtragem Inicial:** Atletas com pontuação total (`total`) igual a zero são removidos da lista.
2.  **Separação por Categoria:** Os atletas são divididos em tabelas separadas, uma para cada categoria (`DivisionClass`).
3.  **Alocação de Grupos:** Utilizando a planilha `DistGrupos.xlsx` como referência, o script atribui um número de grupo para cada atleta.
4.  **Gestão de Eliminados:** Atletas que, devido ao seu ranking, não se encaixam nos grupos principais (marcados com o grupo "29" na planilha de distribuição) são movidos para uma lista de "eliminados".
5.  **Preenchimento de Grupos (BYEs):** O script verifica os grupos que ficaram com menos de 4 competidores e adiciona automaticamente entradas "BYE" para completá-los, garantindo que cada grupo tenha 4 posições.
6.  **Geração dos Combates:** Por fim, para cada grupo completo, o sistema gera as três rodadas de confrontos (MATCH 1, 2 e 3), pareando os atletas de acordo com a formatação padrão do Robin Round.

## Saídas (Outputs)

Após o processamento, a aplicação disponibiliza três arquivos Excel para download:

1.  **Planilha de Grupos (`<nome_da_etapa>.xlsx`):**
    *   Um arquivo Excel com múltiplas abas, onde cada aba corresponde a uma categoria (`DivisionClass`).
    *   Conteúdo: A composição final de cada grupo, incluindo os atletas e os "BYEs" adicionados.

2.  **Planilha de Combates (`<nome_da_etapa>_combates.xlsx`):**
    *   Também é um arquivo com múltiplas abas por categoria, além de uma aba "Total" que consolida todos os confrontos.
    *   Conteúdo: A agenda detalhada dos combates (MATCH 1, 2, 3) para cada grupo, especificando os oponentes.

3.  **Planilha de Eliminados (`eliminados.xlsx`):**
    *   Uma lista simples contendo os dados dos atletas que não foram alocados em nenhum grupo de combate.

## Ferramentas e Tecnologias

*   **Python:** Linguagem de programação principal.
*   **Streamlit:** Framework utilizado para construir a interface web interativa da aplicação.
*   **Pandas:** Biblioteca utilizada para toda a manipulação de dados, leitura e escrita dos arquivos Excel.

## Como Executar o Projeto Localmente

1.  Clone o repositório para a sua máquina.
2.  Abra um terminal na pasta do projeto.
3.  Instale as dependências necessárias:
    ```shell
    pip install -r requirements.txt
    ```
4.  Execute a aplicação Streamlit:
    ```shell
    streamlit run app.py
    ```
5.  Uma nova aba será aberta no seu navegador com a aplicação em funcionamento.
