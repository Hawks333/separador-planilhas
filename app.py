import streamlit as st
import pandas as pd
import os
import zipfile
from io import BytesIO

st.set_page_config(page_title="Separador de Planilhas", layout="centered")

st.title("📊 Separador de Planilha por Coluna A")

st.markdown("""
Envie um arquivo Excel `.xlsx` com a **primeira linha como cabeçalho** e os dados que você deseja separar na **coluna A** (primeira coluna).
O sistema criará arquivos separados com base em cada valor único encontrado nessa coluna.
""")

uploaded_file = st.file_uploader("📁 Envie seu arquivo .xlsx", type=["xlsx"])

if uploaded_file:
    try:
        # Lê o Excel
        df = pd.read_excel(uploaded_file)
        coluna_A = df.columns[0]

        st.success(f"Arquivo carregado com sucesso! Coluna usada para separação: **{coluna_A}**")
        st.write(df.head())

        if st.button("🚀 Separar e baixar arquivos"):
            # Cria os arquivos em memória
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for valor in df[coluna_A].dropna().unique():
                    df_filtrado = df[df[coluna_A] == valor]
                    nome_arquivo = str(valor).strip().replace('/', '_').replace('\\', '_').replace(':', '-')
                    excel_bytes = BytesIO()
                    df_filtrado.to_excel(excel_bytes, index=False, engine='openpyxl')
                    excel_bytes.seek(0)
                    zip_file.writestr(f"{nome_arquivo}.xlsx", excel_bytes.read())

            zip_buffer.seek(0)
            st.download_button(
                label="📦 Baixar arquivos separados (.zip)",
                data=zip_buffer,
                file_name="arquivos_separados.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
