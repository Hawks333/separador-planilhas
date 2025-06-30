import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

st.set_page_config(page_title="Separador de Planilhas", layout="centered")

st.title("📊 Separador de Planilha por Coluna")

st.markdown("""
Envie um arquivo Excel `.xlsx` com a **primeira linha como cabeçalho** e selecione a coluna que deseja usar para separar os dados.
O sistema criará arquivos separados com base em cada valor único encontrado na coluna selecionada.
""")

uploaded_file = st.file_uploader("📁 Envie seu arquivo .xlsx", type=["xlsx"])

if uploaded_file:
    try:
        # Lê o Excel
        df = pd.read_excel(uploaded_file)
        
        # Mostra as colunas disponíveis para seleção
        coluna_separadora = st.selectbox(
            "Selecione a coluna para separar os arquivos:",
            options=df.columns,
            index=0  # Seleciona a primeira coluna por padrão
        )
        
        st.success(f"Arquivo carregado com sucesso! Coluna selecionada para separação: **{coluna_separadora}**")
        st.write("Visualização dos dados (5 primeiras linhas):")
        st.write(df.head())

        if st.button("🚀 Separar e baixar arquivos"):
            # Cria uma coluna temporária com os valores normalizados (tudo em minúsculo)
            df['temp_normalized'] = df[coluna_separadora].astype(str).str.strip().str.lower()
            
            # Cria os arquivos em memória
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                # Agrupa pelos valores normalizados
                for valor_normalizado in df['temp_normalized'].dropna().unique():
                    # Pega o primeiro valor original (para manter a formatação original no arquivo)
                    valor_original = df.loc[df['temp_normalized'] == valor_normalizado, coluna_separadora].iloc[0]
                    
                    # Filtra o dataframe
                    df_filtrado = df[df['temp_normalized'] == valor_normalizado]
                    
                    # Remove a coluna temporária antes de salvar
                    df_filtrado = df_filtrado.drop(columns=['temp_normalized'])
                    
                    # Cria nome do arquivo seguro
                    nome_arquivo = str(valor_original).strip().replace('/', '_').replace('\\', '_').replace(':', '-')
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
