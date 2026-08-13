# Gerador de Combates — Robin Round Individual

Aplicação em Streamlit para organizar um Robin Round Individual de competições de arco e flecha, gerar os confrontos, registrar os resultados e produzir a classificação final da prova.

# Parte 1 — Guia rápido para o usuário

## O que o aplicativo faz

O aplicativo transforma os resultados da qualificatória em uma etapa completa de Robin Round:

1. lê e organiza os atletas da qualificatória;
2. permite confirmar quem participará dos combates;
3. distribui os atletas em grupos;
4. gera os confrontos e os arquivos para impressão;
5. recebe os resultados preenchidos;
6. calcula vitórias, médias, bônus e pontuação final;
7. disponibiliza o Excel final e o CSV da competição.

## Fluxo recomendado

### Etapa 1 — Gerar os combates

Na página **Robin Round Individual**:

1. informe o **Nome da Etapa**;
2. informe o **Local da Prova**;
3. carregue o arquivo bruto da qualificatória (`.txt`, `.csv` ou `.xlsx`);
4. confira a lista de atletas apresentada;
5. desmarque os atletas que não participarão do Robin Round;
6. clique em **Confirmar seleção**;
7. clique em **Processar Combates**;
8. baixe as planilhas e, se necessário, gere os PDFs para impressão.

Ao final dessa etapa, o aplicativo pode gerar:

- planilha de grupos;
- planilha de combates;
- lista de eliminados pela distribuição;
- PDFs de impressão separados por categoria.

### Etapa 2 — Registrar os resultados

Na página **Consolidação de Resultados**:

1. carregue a planilha de combates gerada na etapa anterior;
2. clique em **Gerar template de resultados**;
3. baixe o template;
4. preencha os resultados nas colunas dos sets e do `Shoot-Off`;
5. carregue o template preenchido;
6. carregue novamente o arquivo bruto da qualificatória;
7. clique em **Gerar arquivo final**;
8. baixe o Excel final e o CSV Robin Round.

> O arquivo bruto da qualificatória é solicitado novamente na consolidação para recuperar dados como ID, sigla, clube e pontuações da qualificatória.

## Como preencher os resultados

O template possui uma aba para cada confronto (`MATCH 1`, `MATCH 2` e `MATCH 3`). Preencha somente as colunas de resultado correspondentes a cada atleta:

- `Set 1_a` a `Set 5_a`;
- `SO_a`;
- `Set 1_b` a `Set 5_b`;
- `SO_b`.

Não altere os nomes dos atletas, as categorias, os grupos ou a estrutura das abas. Esses campos são usados para identificar os participantes e consolidar o resultado.

## Arquivos que o usuário recebe

Depois de gerar os combates:

```text
<etapa>_grupos.xlsx
<etapa>_combates.xlsx
<etapa>_eliminados.xlsx
<etapa>_<categoria>.pdf
```

Depois de consolidar os resultados:

```text
<etapa>_final.xlsx
<nome do arquivo bruto> Robin Round.csv
```

# Parte 2 — Detalhes técnicos

## Estrutura esperada do projeto

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

## Componentes Python

### `app.py`

É o ponto de entrada do Streamlit. Configura a página e disponibiliza as rotas:

- **Início**;
- **Robin Round Individual**;
- **Consolidação de Resultados**.

Execução local:

```bash
streamlit run app.py
```

### `pages/robinround.py`

Implementa a preparação dos combates:

1. leitura do arquivo da qualificatória;
2. normalização de nomes, categorias e pontuações;
3. filtragem da sessão, quando aplicável;
4. mapeamento para a categoria de combate;
5. ranking da qualificatória dentro de cada categoria;
6. seleção dos atletas;
7. distribuição conforme `DistGrupos.xlsx`;
8. inclusão de `BYE` em grupos incompletos;
9. geração dos seis confrontos de cada grupo;
10. exportação das planilhas e PDFs.

### `pages/resultados_rr.py`

Implementa a consolidação:

1. criação do template a partir da aba `Total` da planilha de combates;
2. leitura dos resultados preenchidos;
3. identificação dos vencedores;
4. contagem de vitórias e `Shoot-Offs`;
5. cálculo da média dos combates;
6. aplicação dos bônus;
7. cálculo da pontuação total da prova;
8. ranking final por categoria;
9. geração do Excel final;
10. geração do CSV final.

## Arquivo de entrada da qualificatória

São aceitos os formatos:

- `.txt`;
- `.csv`;
- `.xlsx`.

Arquivos `.txt` e `.csv` devem usar `;` como separador. As colunas reconhecidas incluem:

- identificação: `ID`, `Id`, `WaID`, `Athlete ID` ou `AthleteId`;
- nome: `Nome Completo`, `NomeCompleto`, `NOME` ou `FamilyName` + `GivenName`;
- clube: `Clube`, `CLUBE` ou `Country`;
- sigla: `Sigla`, `SIGLA`, `Noc` ou `Club Code`;
- categoria: `Categoria Quali`, `Categoria`, `Division` + `Class` ou `Category`;
- pontuação: `Score`, `10+X` e `X`;
- sessão: `Session`.

Quando `Division` e `Class` existem, a categoria da qualificatória é formada pela concatenação dos dois campos. A categoria é então convertida pela tabela `DePara Categorias.xlsx`.

Se a coluna `Session` existir, somente os atletas com `Session = 1` entram na lista de combates.

## Normalização e ranking da qualificatória

Antes do processamento, a aplicação:

- remove espaços excedentes dos campos textuais;
- converte pontuações numéricas, aceitando vírgula decimal;
- cria o nome completo quando necessário;
- padroniza clube e sigla;
- substitui a categoria original pela categoria de combate quando houver correspondência no de-para.

O ranking é calculado separadamente por categoria de combate, considerando, nesta ordem:

1. `Score`, do maior para o menor;
2. `10+X`, do maior para o menor;
3. `X`, do maior para o menor;
4. nome do atleta como critério técnico de ordenação quando os valores anteriores forem iguais.

Esse ranking é usado para consultar a matriz de distribuição de grupos.

## Distribuição dos grupos

### `DistGrupos.xlsx`

A matriz define o grupo de cada posição do ranking de acordo com a quantidade de atletas da categoria.

- o índice representa o ranking;
- as colunas representam a quantidade de atletas;
- o valor encontrado representa o grupo;
- o valor `99` representa uma posição inválida ou não distribuída.

Atletas associados ao valor `99` são enviados para a planilha de eliminados. Grupos com menos de quatro atletas são completados com participantes `BYE`.

### Confrontos

Cada grupo de quatro posições gera seis confrontos:

| Match | Confronto |
|---|---|
| MATCH 1 | Atleta 1 × Atleta 4 |
| MATCH 1 | Atleta 2 × Atleta 3 |
| MATCH 2 | Atleta 1 × Atleta 3 |
| MATCH 2 | Atleta 2 × Atleta 4 |
| MATCH 3 | Atleta 1 × Atleta 2 |
| MATCH 3 | Atleta 3 × Atleta 4 |

Confrontos formados exclusivamente por `BYE` não são incluídos na planilha de combates.

## Regras para determinar o vencedor

### Categorias de soma

Categorias cujo código começa com `C` usam a soma dos valores dos sets. O maior total vence. Se houver empate, o maior valor de `Shoot-Off` define o vencedor.

### Categorias de sets

Nas demais categorias, cada set vale:

- vitória no set: 2 pontos;
- empate no set: 1 ponto para cada atleta;
- derrota no set: 0 pontos.

O atleta que atingir pelo menos 6 pontos de sets primeiro vence. Se o confronto terminar empatado, o `Shoot-Off` é usado como desempate.

### `BYE`

Quando apenas um dos lados é `BYE`, o atleta real é considerado vencedor automaticamente.

## Cálculo da pontuação final

A pontuação total do atleta é formada pela soma de:

- pontos do Round;
- bônus do Round;
- número de vitórias nos combates;
- bonificação por `Shoot-Off`;
- pontuação do ranking de médias no grupo;
- bônus da média;
- bônus do grupo.

### Arquivos de regras de pontuação

#### `Pontos Round.xlsx`

Converte o ranking da qualificatória em pontos de Round. A busca utiliza a coluna `Rank` e o valor correspondente em `Pontos Round`.

#### `Bonus Round.xlsx`

Define o bônus da qualificatória por:

- categoria de combate;
- gênero;
- faixa de pontuação mínima e máxima.

#### `Bonus Médias.xlsx`

Define o bônus associado à média dos combates por categoria e faixa de média.

#### `Pontos FPAF.xlsx`

Define a pontuação individual usada no CSV final. A busca considera:

- ranking final;
- quantidade de atletas da categoria.

## Cálculo da média dos combates

Para categorias que não começam com `C`, a média é calculada em duas etapas:

1. para cada `Match`, identifica-se quantos sets foram efetivamente disputados usando a mesma regra de comparação set a set e de pontuação do vencedor;
2. somam-se as pontuações do atleta em todos esses sets, inclusive quando algum set válido tem pontuação `0`, e divide-se pela quantidade de sets disputados no combate;
3. calcula-se a média aritmética das médias dos três `Matches` do atleta.

Células vazias representam sets não disputados. O valor `0` preenchido representa um set válido e entra no cálculo.

Dessa forma, cada combate tem o mesmo peso, mesmo quando possui quantidades diferentes de sets — por exemplo, 3, 4 ou 5 sets.

Para categorias iniciadas com `C`, permanece o cálculo baseado na média dos totais de cada combate.

## Ranking das médias e empates

A classificação das médias dentro de cada grupo deve usar ranking competitivo, sem desempate pelo nome.

Exemplos:

- médias diferentes: `1º, 2º, 3º, 4º`;
- empate na segunda posição: `1º, 2º, 2º, 4º`;
- empate na primeira posição: `1º, 1º, 3º, 4º`;
- empate na terceira posição: `1º, 2º, 3º, 3º`.

Os bônus correspondentes são atribuídos à posição competitiva. Portanto, em um empate na terceira posição, os quatro atletas recebem:

```text
2,0 / 1,5 / 1,0 / 1,0
```

A quarta posição é consumida pelo empate e o último atleta fica na quinta posição. Se não houver bônus definido para a quinta posição, ele recebe `0,0`.

## Ranking final

O ranking final é calculado separadamente por categoria, em ordem decrescente de `Pontuação Total da Prova`.

Atletas com a mesma pontuação recebem a mesma posição final pelo ranking competitivo. O nome pode ser usado apenas para ordenar visualmente as linhas, não para alterar a posição atribuída ao empate.

## Excel final

O arquivo final contém uma aba por categoria. As principais colunas são:

```text
Pos Final
Atleta
Cat Round
Cat Combate
Clube
Round 1
Round 2
Total Round
Pontos Round
Bonus Round
Nº de Vitórias Combates
Bonificação Shoot-Offs
Média dos Combates
Ranking Médias no Grupo
Bonus Média
Bonus Grupo
Pontuação Total da Prova
```

## CSV final

O CSV possui as colunas:

```text
RANKING FINAL
ID
NOME
CATEGORIA AGRUPADA
SIGLA
CLUBE
PONTUAÇÃO INDIVIDUAL
```

Regras de preenchimento:

- `RANKING FINAL`: posição final do atleta;
- `ID`: recuperado do arquivo bruto;
- `NOME`: nome do atleta no arquivo final;
- `CATEGORIA AGRUPADA`: nome da aba/categoria do Excel final;
- `SIGLA`: recuperada do arquivo bruto por nome ou clube;
- `CLUBE`: utiliza o valor final, com fallback para o arquivo bruto;
- `PONTUAÇÃO INDIVIDUAL`: consultada em `Pontos FPAF.xlsx`.

Atletas `BYE` não são incluídos no CSV nem na contagem de atletas da categoria. O arquivo usa `;` como separador e codificação UTF-8 com BOM.

## Arquivos auxiliares e imagens

Os seguintes arquivos devem permanecer disponíveis na raiz do projeto, com os nomes exatamente iguais:

- `DistGrupos.xlsx`;
- `DePara Categorias.xlsx`;
- `Pontos Round.xlsx`;
- `Pontos FPAF.xlsx`;
- `Bonus Round.xlsx`;
- `Bonus Médias.xlsx`;
- `Fundo Robin Round Individual - Sets.png`;
- `Fundo Robin Round Individual - Soma.png`.

Categorias iniciadas com `C` usam o fundo de soma nos PDFs. As demais usam o fundo de sets.

## Dependências e execução local

Dependências principais:

```text
streamlit
pandas
openpyxl
fpdf
Pillow
```

Instalação:

```bash
pip install -r requirements.txt
```

Execução:

```bash
streamlit run app.py
```

O ambiente Python recomendado está definido em `runtime.txt`.

## Implantação no Streamlit Community Cloud

Para publicar o projeto:

1. envie `app.py`, a pasta `pages`, `requirements.txt` e `runtime.txt` para o repositório;
2. mantenha os arquivos `.xlsx` e `.png` na raiz;
3. configure `app.py` como arquivo principal;
4. confirme que os nomes dos arquivos auxiliares não foram alterados.

## Solução de problemas

- **Categoria não reconhecida:** confira `DePara Categorias.xlsx` e os campos `Division`/`Class` do arquivo de entrada.
- **Poucos atletas na lista:** verifique o filtro `Session = 1` e a seleção manual.
- **Atletas eliminados:** confira a matriz `DistGrupos.xlsx`; o valor `99` envia o atleta para a lista de eliminados.
- **Resultado não identificado:** preserve os nomes dos atletas e as abas criadas no template.
- **CSV sem pontuação individual:** confira a existência de `Pontos FPAF.xlsx` e se a combinação entre ranking e quantidade de atletas está cadastrada.
- **Erro ao gerar PDF:** confirme a presença dos dois arquivos PNG na raiz do projeto.
