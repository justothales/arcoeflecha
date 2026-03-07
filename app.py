import streamlit as st
from pages.robinround import show_page as show_robin_round_page

# Configuração global da página
st.set_page_config(
    page_title="Gerador de Combates",
    page_icon="🎯",
    layout="wide"
)

# Função que define o conteúdo da página inicial
def show_inicio_page():
    """Renderiza o conteúdo da página inicial."""
    st.title("Gerador de Combates de Arco e Flecha")
    st.markdown("""
    ### Bem-vindo ao Gerador de Combates!

    Esta aplicação centraliza diversas ferramentas para auxiliar na organização de competições de tiro com arco.

    **Para começar, selecione a funcionalidade desejada no menu da barra lateral à esquerda.**
    """)

# Define a estrutura de navegação
pg = st.navigation([
    st.Page(show_inicio_page, title="Início", default=True, icon="🏠"),
    st.Page(show_robin_round_page, title="Robin Round", icon="🎯")
])

# Executa a página selecionada
pg.run()
