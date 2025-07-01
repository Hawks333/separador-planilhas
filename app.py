if uploaded_file:
    try:
        # Lê o Excel
        df = pd.read_excel(uploaded_file)

        # Remove colunas completamente vazias
        df = df.dropna(axis=1, how='all')

        # Remove colunas com nomes "Unnamed: X"
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

        # Mostra as colunas disponíveis para seleção
        coluna_separadora = st.selectbox(
            "Selecione a coluna para separar os arquivos:",
            options=df.columns,
            index=0
        )

        st.success(f"Arquivo carregado com sucesso! Coluna selecionada para separação: **{coluna_separadora}**")
        st.write("Visualização dos dados (5 primeiras linhas):")
        st.write(df.head())

        if st.button("🚀 Separar e baixar arquivos"):
            df['temp_normalized'] = df[coluna_separadora].astype(str).str.strip().str.lower()
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for valor_normalizado in df['temp_normalized'].dropna().unique():
                    valor_original = df.loc[df['temp_normalized'] == valor_normalizado, coluna_separadora].iloc[0]
                    df_filtrado = df[df['temp_normalized'] == valor_normalizado]
                    df_filtrado = df_filtrado.drop(columns=['temp_normalized'])
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
