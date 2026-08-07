# Gerador de Combates — Robin Round Individual

Aplicação em Streamlit para organizar a fase de combates de um Robin Round Individual de competições de arco e flecha, consolidar os resultados e gerar os arquivos finais da prova.

## Funcionalidades

A aplicação permite:

- carregar resultados de uma fase qualificatória;
- aceitar arquivos `.txt`, `.csv` e `.xlsx`;
- normalizar nomes, clubes, categorias e pontuações;
- filtrar atletas da `Session = 1`, quando essa coluna existir;
- mapear categorias da qualificatória para categorias agrupadas de combate;
- calcular o ranking dentro de cada categoria;
- selecionar manualmente os atletas que avançam para os combates;
- distribuir atletas nos grupos conforme a matriz de distribuição;
- preencher grupos incompletos com participantes `BYE`;
- gerar a planilha de grupos;
- gerar a planilha de combates;
- gerar a lista de atletas eliminados pela distribuição;
- gerar PDFs prontos para impressão;
- criar um template para lançamento dos resultados dos combates;
- consolidar vencedores, vitórias, médias, bônus e pontuação final;
- gerar o arquivo final em Excel;
- gerar um CSV final no formato utilizado pela federação.

## Estrutura do projeto

```text
projeto/
├── app.py
├── requirements.txt
├── runtime.txt
├── README.md
├── DistGrupos.xlsx
├── DePara Categorias.xlsx
├── Pontos Round.xlsx
├── Pontos FPAF.xlsx
├── Bonus Round.xlsx
├── Bonus Médias.xlsx
├── Fundo Robin Round Individual - Sets.png
├── Fundo Robin Round Individual - Soma.png
└── pages/
    ├── __init__.py
    ├── robinround.py
    └── resultados_rr.py
```

## Arquivos Python

### `app.py`

É o ponto de entrada da aplicação. Configura o Streamlit e disponibiliza as páginas:

- **Início**;
- **Robin Round Individual**;
- **Consolidação de Resultados**.

Execute a aplicação com:

```bash
streamlit run app.py
```

### `pages/robinround.py`

Responsável pela primeira etapa do fluxo:

1. leitura dos resultados da qualificatória;
2. normalização dos dados dos atletas;
3. conversão das categorias;
4. cálculo do ranking;
5. seleção dos atletas que avançam;
6. distribuição nos grupos;
7. inclusão de `BYE` quando necessário;
8. criação dos confrontos;
9. exportação das planilhas;
10. geração dos PDFs para impressão.

### `pages/resultados_rr.py`

Responsável pela consolidação dos resultados:

1. criação do template de resultados;
2. leitura dos resultados preenchidos;
3. identificação dos vencedores;
4. cálculo das vitórias;
5. cálculo das médias dos combates;
6. aplicação dos bônus;
7. cálculo da pontuação total da prova;
8. classificação final por categoria;
9. geração do Excel final;
10. geração do CSV final do Robin Round.

## Arquivos auxiliares

### `DistGrupos.xlsx`

Define a distribuição de grupos de acordo com a quantidade de atletas e o ranking.

O valor `99` representa uma combinação inválida ou uma posição que não deve participar da distribuição. Esses atletas são enviados para a lista de eliminados.

### `DePara Categorias.xlsx`

Mapeia a categoria original da qualificatória para a categoria agrupada utilizada nos combates.

Exemplos:

```text
B65M  -> B50M
CS15M -> C40M
R50F  -> R60M
W1F   -> W150M
```

### `Pontos Round.xlsx`

Converte o ranking da qualificatória em pontos de Round.

### `Pontos FPAF.xlsx`

Define a pontuação individual do resultado final.

- a primeira coluna contém o ranking final;
- as colunas seguintes representam a quantidade de atletas da categoria;
- o valor da pontuação é obtido no cruzamento entre o ranking final e a quantidade de atletas da categoria.

### `Bonus Round.xlsx`

Contém os bônus relacionados à pontuação da fase qualificatória, considerando a categoria, o gênero e a faixa de pontuação.

### `Bonus Médias.xlsx`

Contém os bônus relacionados à média dos combates, considerando a categoria e a faixa de pontuação.

### Arquivos PNG

Os arquivos abaixo são utilizados como fundo dos PDFs de impressão:

```text
Fundo Robin Round Individual - Sets.png
Fundo Robin Round Individual - Soma.png
```

Categorias iniciadas com `C` utilizam o modelo de soma. As demais utilizam o modelo de sets.

## Entrada da qualificatória

A aplicação aceita arquivos nos formatos:

- `.txt`;
- `.csv`;
- `.xlsx`.

Arquivos `.txt` e `.csv` devem utilizar `;` como separador. A aplicação tenta ler arquivos com as codificações:

1. UTF-8 com BOM;
2. `cp1252`;
3. `latin1`.

O arquivo pode conter, entre outras, as seguintes colunas:

- `Nome Completo` ou `NOME`;
- `FamilyName` e `GivenName`;
- `ID`, `WaID` ou identificador equivalente;
- `Clube` ou `CLUBE`;
- `Sigla`, `SIGLA`, `Noc` ou `Club Code`;
- `Division`;
- `Class`;
- `Score`;
- `10+X`;
- `X`;
- `Session`.

Quando `Division` e `Class` estiverem presentes, a categoria da qualificatória é formada pela combinação desses dois campos.

## Fluxo de uso

### 1. Gerar os combates

Na página **Robin Round Individual**:

1. informe o nome da etapa;
2. informe o local da prova;
3. carregue o arquivo bruto da qualificatória;
4. confira os atletas carregados;
5. desmarque os atletas que não participarão;
6. confirme a seleção;
7. clique em **Processar Combates**;
8. baixe os arquivos gerados.

### 2. Preencher os resultados

Na página **Consolidação de Resultados**:

1. carregue a planilha de combates;
2. gere o template de resultados;
3. preencha os resultados nas colunas dos sets e do `Shoot-Off`;
4. carregue o template preenchido;
5. carregue novamente o arquivo bruto da qualificatória;
6. clique em **Gerar arquivo final**.

## Combates gerados

Cada grupo de quatro atletas gera seis confrontos:

| Match | Confronto |
|---|---|
| MATCH 1 | Atleta 1 × Atleta 4 |
| MATCH 1 | Atleta 2 × Atleta 3 |
| MATCH 2 | Atleta 1 × Atleta 3 |
| MATCH 2 | Atleta 2 × Atleta 4 |
| MATCH 3 | Atleta 1 × Atleta 2 |
| MATCH 3 | Atleta 3 × Atleta 4 |

Os confrontos são organizados por categoria e grupo.

## Regras de resultado

### Categorias de soma

Para categorias iniciadas com `C`, o vencedor é definido pela soma das pontuações dos sets. Em caso de empate, o `Shoot-Off` é utilizado como desempate.

### Categorias de sets

Para as demais categorias:

- vitória em um set: 2 pontos;
- empate em um set: 1 ponto para cada atleta;
- derrota em um set: 0 pontos.

O primeiro atleta a atingir pelo menos 6 pontos de set é considerado vencedor. Se o confronto terminar empatado, o `Shoot-Off` é utilizado como desempate.

### `BYE`

Quando um confronto possui um atleta real e um `BYE`, o atleta real é considerado vencedor automaticamente.

## Arquivos gerados

Após o processamento dos combates, a aplicação disponibiliza:

```text
<etapa>_grupos.xlsx
<etapa>_combates.xlsx
<etapa>_eliminados.xlsx
```

Também podem ser gerados PDFs separados por categoria:

```text
<etapa>_<categoria>.pdf
```

Após a consolidação dos resultados, são gerados:

```text
<etapa>_final.xlsx
<nome do arquivo bruto> Robin Round.csv
```

## CSV final do Robin Round

O CSV final possui exatamente estas colunas:

```text
RANKING FINAL
ID
NOME
CATEGORIA AGRUPADA
SIGLA
CLUBE
PONTUAÇÃO INDIVIDUAL
```

### Regras do CSV

- `RANKING FINAL`: utiliza a posição final do arquivo consolidado;
- `ID`: é recuperado do arquivo bruto da qualificatória;
- `NOME`: utiliza o nome completo do atleta;
- `CATEGORIA AGRUPADA`: utiliza a categoria de combate;
- `SIGLA`: é recuperada do arquivo bruto por correspondência de nome ou clube;
- `CLUBE`: utiliza o clube do arquivo final, com fallback para o arquivo bruto;
- `PONTUAÇÃO INDIVIDUAL`: é calculada usando `Pontos FPAF.xlsx`.

Os atletas `BYE` não são incluídos no CSV e não entram na contagem de atletas da categoria.

O CSV é gerado com:

- separador `;`;
- codificação UTF-8 com BOM;
- suporte a caracteres acentuados.

## Dependências

As principais dependências são:

```text
streamlit
pandas
openpyxl
fpdf
Pillow
```

Para instalar:

```bash
pip install -r requirements.txt
```

## Execução local

1. Abra o terminal na pasta do projeto.
2. Instale as dependências:

```bash
pip install -r requirements.txt
```

3. Execute o Streamlit:

```bash
streamlit run app.py
```

4. Abra no navegador o endereço informado pelo Streamlit.

## Implantação no Streamlit Community Cloud

Para publicar o projeto:

1. envie todos os arquivos para o repositório;
2. confirme que `app.py` está na raiz;
3. confirme que a pasta `pages` está na raiz;
4. mantenha os arquivos `.xlsx` e `.png` na raiz;
5. configure `app.py` como arquivo principal;
6. mantenha `requirements.txt` e `runtime.txt` no repositório.

## Observações

- Os nomes dos arquivos auxiliares devem ser mantidos exatamente, incluindo espaços, acentos e extensão.
- O arquivo `Pontos FPAF.xlsx` é necessário para gerar corretamente a pontuação individual do CSV final.
- O arquivo bruto deve ser fornecido novamente na etapa de consolidação para recuperar o ID e a sigla dos atletas.
- O preenchimento dos resultados deve ser feito no template gerado pela própria aplicação.
- Em caso de erro, confira principalmente os nomes das colunas do arquivo bruto e a presença dos arquivos auxiliares na raiz do projeto.
