import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font

# Definir as colunas padrão
DEFAULT_COLUMNS = [
    "Sample", "Pluton", "Group", "Rock_type", "Observation", "Tectonic_setting", "Location_notes",
    "Age", "Reference", "Colour", "Symbol", "Size", "SiO2", "TiO2", "Al2O3", "FeO", "FeOt",
    "Fe2O3", "Fe2O3t", "MnO", "MgO", "CaO", "K2O", "Na2O", "P2O5", "Total", "H2Ot", "LOI",
    "Li", "Be", "B", "Sc", "V", "Cr", "Ni", "Cu", "Zn", "Rb", "Sr", "Y", "Zr", "Nb", "Cs", 
    "Ba", "La", "Ce", "Pr", "Nd", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb", 
    "Lu", "Hf", "Ta", "Pb", "Th", "U", "Co", "Mo", "W", "Ga", "Ge", "As", "In", "Sn", 
    "Sb", "Cd"
]

st.title("Reorganizador de Colunas para Dados Geoquímicos")

# Adicionar assinatura no rodapé
st.markdown("---")  # Linha separadora
st.markdown(
    """
    **Desenvolvido por [Pedro Armond](https://www.researchgate.net/profile/Pedro-Armond)**  
    📧 E-mail: [pedro.armond@aluno.ufop.edu.br](mailto:pedro.armond@aluno.ufop.edu.br)  
    🌐 ResearchGate: [https://www.researchgate.net/profile/Pedro-Armond](https://www.researchgate.net/profile/Pedro-Armond)
    """
)

# Upload da planilha
uploaded_file = st.file_uploader("Carregue sua planilha (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Exibir as abas disponíveis no arquivo Excel
    workbook = pd.ExcelFile(uploaded_file)
    sheet_names = workbook.sheet_names
    st.write("A planilha contém as seguintes abas:")
    selected_sheet = st.selectbox("Selecione a aba que contém os dados:", sheet_names)

    # Leitura da aba selecionada
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    st.write("Pré-visualização dos dados:")
    st.dataframe(df)

    # Opção para selecionar a ordem das colunas
    st.subheader("Selecione a ordem das colunas:")
    selected_columns = st.multiselect(
        "Escolha as colunas desejadas na ordem correta:",
        DEFAULT_COLUMNS,
        default=[]
    )

    # Botão para selecionar todas as colunas automaticamente
    if st.button("Marcar todas as colunas (recomendado)"):
        selected_columns = DEFAULT_COLUMNS

    st.write("Ordem selecionada:")
    st.write(selected_columns)

    # Reorganizar colunas ao clicar no botão "Salvar"
    if st.button("Salvar Arquivo"):
        # Identificar colunas faltantes e excedentes
        missing_columns = list(set(DEFAULT_COLUMNS) - set(df.columns))
        extra_columns = list(set(df.columns) - set(DEFAULT_COLUMNS))
        
        # Adicionar colunas faltantes com valores NaN
        for col in missing_columns:
            df[col] = None

        # Reordenar colunas
        reordered_columns = DEFAULT_COLUMNS + extra_columns
        df = df[reordered_columns]

        # Salvar arquivo
        original_name = uploaded_file.name
        base_name, ext = os.path.splitext(original_name)
        new_file_name = f"{base_name}_modified{ext}"
        df.to_excel(new_file_name, index=False, engine="openpyxl")

        st.success(f"Arquivo salvo como {new_file_name}.")
        st.write("Baixe o arquivo abaixo:")
        with open(new_file_name, "rb") as file:
            st.download_button(
                label="📥 Baixar Arquivo Modificado",
                data=file,
                file_name=new_file_name
            )
