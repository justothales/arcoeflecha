import pandas as pd
import streamlit as st
import io

# Helper function to convert dataframe to excel bytes
def to_excel(dfs, multi_sheet=False):
    """
    Converts a pandas DataFrame (or a dictionary of DataFrames for multi-sheet)
    into an in-memory Excel file (bytes).
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if multi_sheet:
            for sheet_name, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            dfs.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def processar_combates(df, df_dist_grupos, etapa, local_da_prova):
    """
    Core logic from the Jupyter Notebook to process the combat data.
    Takes dataframes and user inputs, returns three dataframes for output.
    """
    # 1 - Retirar todas as linhas cuja coluna "total" seja igual a zero
    df = df[df['total'] != 0]

    # 2 - Separar em tabelas diferentes para cada código DivisionClass
    tabelas_division = {}
    for division_class, data in df.groupby('DivisionClass'):
        tabelas_division[division_class] = data.copy()

    tabela_eliminados = pd.DataFrame()
    combates_geral_df = pd.DataFrame()
    
    final_grupos_tabelas = {}

    for division_class, tabela in tabelas_division.items():
        num_pessoas = len(tabela)
        if num_pessoas == 0:
            continue

        valores_grupo = []
        for _, row in tabela.iterrows():
            rank = row['Rank']
            if num_pessoas in df_dist_grupos.columns and rank in df_dist_grupos.index:
                valor_grupo = df_dist_grupos[num_pessoas][rank]
                valores_grupo.append(valor_grupo)
            else:
                valores_grupo.append(29) # Default to eliminated if not in dist_grupos

        tabela['grupo'] = valores_grupo
        tabela['pos_grupo'] = tabela.groupby('grupo').cumcount() + 1
        
        eliminados = tabela[tabela['grupo'] == 29]
        tabela_eliminados = pd.concat([tabela_eliminados, eliminados])
        
        tabela = tabela[tabela['grupo'] != 29]
        if tabela.empty:
            continue

        grupos_faltantes = tabela['grupo'].value_counts()
        grupos_completar = grupos_faltantes[grupos_faltantes < 4].index

        for grupo in grupos_completar:
            num_linhas_grupo = len(tabela[tabela['grupo'] == grupo])
            num_linhas_preencher = 4 - num_linhas_grupo
            
            for i in range(num_linhas_preencher):
                posicao_grupo = num_linhas_grupo + 1 + i
                linha_preenchimento = {
                    'prova': tabela['prova'].iloc[0], 'DivisionClass': division_class,
                    'WaID': 0, 'Rank': 0, 'NomeCompleto': f'BYE {grupo}{posicao_grupo}',
                    'local_prova': f'BYE {grupo}{posicao_grupo}', 'D2_Score': grupo, 'D1_Score': grupo,
                    'total': grupo * 2, 'Score_10': 0, 'Score_9': 0, '10+X': 0, 'X': 0,
                    'grupo': grupo, 'pos_grupo': posicao_grupo
                }
                tabela = pd.concat([tabela, pd.DataFrame([linha_preenchimento])], ignore_index=True)

        tabela.sort_values(by=['grupo', 'pos_grupo'], inplace=True)
        tabela.reset_index(drop=True, inplace=True)
        final_grupos_tabelas[division_class] = tabela

        combates_divisao_df = pd.DataFrame()
        for _, grupo_df in tabela.groupby('grupo'):
            grupo_df.reset_index(drop=True, inplace=True)
            combates_data = []
            
            # MATCH 1
            combates_data.append({'match': 'MATCH 1', 'genero': division_class, 'GRUPO': f'GRUPO {grupo_df.iloc[0]["grupo"]}', 'nome a': grupo_df.iloc[0]['NomeCompleto'], 'clube a': grupo_df.iloc[0]['local_prova'], 'rank a': grupo_df.iloc[0]['Rank'], 'nome b': grupo_df.iloc[3]['NomeCompleto'], 'clube b': grupo_df.iloc[3]['local_prova'], 'rank b': grupo_df.iloc[3]['Rank'], 'ETAPA': etapa, 'LOCAL': local_da_prova})
            combates_data.append({'match': 'MATCH 1', 'genero': division_class, 'GRUPO': f'GRUPO {grupo_df.iloc[1]["grupo"]}', 'nome a': grupo_df.iloc[1]['NomeCompleto'], 'clube a': grupo_df.iloc[1]['local_prova'], 'rank a': grupo_df.iloc[1]['Rank'], 'nome b': grupo_df.iloc[2]['NomeCompleto'], 'clube b': grupo_df.iloc[2]['local_prova'], 'rank b': grupo_df.iloc[2]['Rank'], 'ETAPA': etapa, 'LOCAL': local_da_prova})
            # MATCH 2
            combates_data.append({'match': 'MATCH 2', 'genero': division_class, 'GRUPO': f'GRUPO {grupo_df.iloc[0]["grupo"]}', 'nome a': grupo_df.iloc[0]['NomeCompleto'], 'clube a': grupo_df.iloc[0]['local_prova'], 'rank a': grupo_df.iloc[0]['Rank'], 'nome b': grupo_df.iloc[2]['NomeCompleto'], 'clube b': grupo_df.iloc[2]['local_prova'], 'rank b': grupo_df.iloc[2]['Rank'], 'ETAPA': etapa, 'LOCAL': local_da_prova})
            combates_data.append({'match': 'MATCH 2', 'genero': division_class, 'GRUPO': f'GRUPO {grupo_df.iloc[1]["grupo"]}', 'nome a': grupo_df.iloc[1]['NomeCompleto'], 'clube a': grupo_df.iloc[1]['local_prova'], 'rank a': grupo_df.iloc[1]['Rank'], 'nome b': grupo_df.iloc[3]['NomeCompleto'], 'clube b': grupo_df.iloc[3]['local_prova'], 'rank b': grupo_df.iloc[3]['Rank'], 'ETAPA': etapa, 'LOCAL': local_da_prova})
            # MATCH 3
            combates_data.append({'match': 'MATCH 3', 'genero': division_class, 'GRUPO': f'GRUPO {grupo_df.iloc[0]["grupo"]}', 'nome a': grupo_df.iloc[0]['NomeCompleto'], 'clube a': grupo_df.iloc[0]['local_prova'], 'rank a': grupo_df.iloc[0]['Rank'], 'nome b': grupo_df.iloc[1]['NomeCompleto'], 'clube b': grupo_df.iloc[1]['local_prova'], 'rank b': grupo_df.iloc[1]['Rank'], 'ETAPA': etapa, 'LOCAL': local_da_prova})
            combates_data.append({'match': 'MATCH 3', 'genero': division_class, 'GRUPO': f'GRUPO {grupo_df.iloc[2]["grupo"]}', 'nome a': grupo_df.iloc[2]['NomeCompleto'], 'clube a': grupo_df.iloc[2]['local_prova'], 'rank a': grupo_df.iloc[2]['Rank'], 'nome b': grupo_df.iloc[3]['NomeCompleto'], 'clube b': grupo_df.iloc[3]['local_prova'], 'rank b': grupo_df.iloc[3]['Rank'], 'ETAPA': etapa, 'LOCAL': local_da_prova})

            combates_divisao_df = pd.concat([combates_divisao_df, pd.DataFrame(combates_data)], ignore_index=True)
        
        combates_geral_df = pd.concat([combates_geral_df, combates_divisao_df], ignore_index=True)

    # Prepare combat sheets for Excel output
    combates_para_salvar = {}
    if not combates_geral_df.empty:
        # Filter out BYE matches
        combates_sem_bye = combates_geral_df[~combates_geral_df['nome a'].str.contains("BYE", na=False)]
        combates_sem_bye = combates_sem_bye[~combates_sem_bye['nome b'].str.contains("BYE", na=False)]
        
        combates_para_salvar['Total'] = combates_sem_bye
        for division in combates_sem_bye['genero'].unique():
            combates_para_salvar[division] = combates_sem_bye[combates_sem_bye['genero'] == division]

    return final_grupos_tabelas, combates_para_salvar, tabela_eliminados

# --- Streamlit UI ---
st.set_page_config(layout="centered")
st.title("Gerador de Combates - Robin Round Individual")

st.markdown("""
Esta ferramenta automatiza a criação de planilhas de combates a partir de uma planilha de resultados.

**Instruções:**
1.  Preencha o nome da **Etapa** e o **Local da Prova**.
2.  Faça o upload da planilha de **Resultados da Prova** (formato `.xlsx`).
3.  Faça o upload da planilha de **Distribuição de Grupos** (`DistGrupos.xlsx`).
4.  Clique no botão **"Gerar Combates"**.
5.  Aguarde o processamento e faça o download dos arquivos gerados.
""")

st.header("1. Informações da Prova")
etapa_input = st.text_input("Nome da Etapa", "7º Outdoor FPAF - 4º Robin Round Individual")
local_input = st.text_input("Local da Prova", "Mairiporã")

st.header("2. Upload de Arquivos")
uploaded_file = st.file_uploader("Carregue a planilha de Resultados da Prova (.xlsx)", type=["xlsx"])
dist_grupos_file = st.file_uploader("Carregue a planilha de Distribuição de Grupos (`DistGrupos.xlsx`)", type=["xlsx"])

if st.button("Gerar Combates"):
    if uploaded_file is not None and dist_grupos_file is not None:
        try:
            with st.spinner('Processando... Por favor, aguarde.'):
                df_prova = pd.read_excel(uploaded_file)
                df_dist_grupos = pd.read_excel(dist_grupos_file, index_col=0)

                # Process data
                grupos_final, combates_final, eliminados_final = processar_combates(
                    df_prova, df_dist_grupos, etapa_input, local_input
                )

            st.success("Processamento concluído com sucesso!")

            # Convert dataframes to Excel bytes
            if grupos_final:
                grupos_xlsx = to_excel(grupos_final, multi_sheet=True)
                st.download_button(
                    label="⬇️ Baixar Planilha de Grupos",
                    data=grupos_xlsx,
                    file_name=f"{etapa_input}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if combates_final:
                combates_xlsx = to_excel(combates_final, multi_sheet=True)
                st.download_button(
                    label="⬇️ Baixar Planilha de Combates",
                    data=combates_xlsx,
                    file_name=f"{etapa_input}_combates.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if not eliminados_final.empty:
                eliminados_xlsx = to_excel(eliminados_final)
                st.download_button(
                    label="⬇️ Baixar Planilha de Eliminados",
                    data=eliminados_xlsx,
                    file_name="eliminados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Ocorreu um erro durante o processamento: {e}")
            st.warning("Verifique se as colunas das planilhas estão corretas e tente novamente.")

    else:
        st.warning("Por favor, faça o upload dos dois arquivos necessários antes de gerar os combates.")
