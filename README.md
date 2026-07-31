# Gerador de Combates para Competições de Arco e Flecha (Robin Round)

Esta aplicação Streamlit gera os combates de um Robin Round Individual a partir dos resultados de uma fase qualificatória.

## O que o projeto faz

* Carrega resultados de atletas em `.txt`, `.csv` ou `.xlsx`.
* Normaliza dados e identifica categorias de combate com base em `DePara Categorias.xlsx`.
* Calcula o rank por categoria e aloca atletas em grupos usando `DistGrupos.xlsx`.
* Preenche grupos de 4 atletas com entradas `BYE` quando necessário.
* Gera a agenda de combates por categoria.
* Produz arquivos Excel e PDFs prontos para impressão.

## Entradas necessárias

### Upload do usuário

A interface aceita um arquivo de resultados em um destes formatos:

* `.txt`
* `.csv`
* `.xlsx`

O arquivo deve conter as informações dos atletas, incluindo pelo menos nome completo e categoria.

* Se a coluna `Session` existir, apenas atletas com `Session = 1` serão considerados para seleção.

### Arquivos obrigatórios na raiz do projeto

Coloque estes arquivos junto ao `app.py`:

* `DistGrupos.xlsx` — distribuição de grupos por número de atletas e rank.
* `DePara Categorias.xlsx` — mapeamento de `Categoria Quali` para `Categoria Combates`.
* `Fundo Robin Round Individual - Sets.png` — imagem de fundo usada para categorias de sets.
* `Fundo Robin Round Individual - Soma.png` — imagem de fundo usada para categorias de soma.

## Saídas geradas

Após processar, o app disponibiliza para download:

* `*<nome_da_etapa>*_grupos.xlsx` — planilha com grupos por categoria.
* `*<nome_da_etapa>*_combates.xlsx` — planilha com todos os combates e abas por categoria.
* `*<nome_da_etapa>*_eliminados.xlsx` — lista de atletas eliminados.

### PDFs de impressão

Ao gerar os PDFs de impressão, o app produz um arquivo PDF por categoria de combate. Cada PDF contém todas as páginas de combates daquela categoria, geradas com o fundo PNG correto.

## Estrutura mínima de arquivos

### Na raiz do projeto

* `app.py`
* `requirements.txt`
* `runtime.txt` (para deploy em Streamlit Cloud)
* `DistGrupos.xlsx`
* `DePara Categorias.xlsx`
* `Fundo Robin Round Individual - Sets.png`
* `Fundo Robin Round Individual - Soma.png`

### Na pasta `pages`

* `pages/robinround.py`

## Dependências

As dependências necessárias incluem:

* `streamlit`
* `pandas`
* `openpyxl`
* `fpdf`
* `Pillow`

> Observação: o fluxo de PDF usa `fpdf` e imagens PNG como fundo. A conversão DOCX → PDF não é mais necessária para o processo de impressão atual.

## Como executar localmente

1. Abra o terminal na pasta do projeto.
2. Instale as dependências:

```bash
pip install -r requirements.txt
```

3. Execute o Streamlit:

```bash
streamlit run app.py
```

4. Acesse o app no navegador quando o Streamlit indicar o endereço.

## Notas importantes

* Garanta que as imagens de fundo PNG estejam presentes na raiz do projeto.
* O app carrega automaticamente `DistGrupos.xlsx` e `DePara Categorias.xlsx`.
* Os PDFs são gerados por categoria de combate, com todas as páginas da mesma categoria no mesmo arquivo.
* Se precisar suportar arquivos `.xls`, mantenha `xlrd` instalado, mas não é obrigatório para o fluxo principal.
