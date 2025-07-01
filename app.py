import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from openpyxl import load_workbook, Workbook

st.set_page_config(page_title="Separador de Planilhas com Formatação", layout="centered")

st.title("📊 Separador de Planilha com Formatação")

st.markdown("""
Envie um arquivo Excel `.xlsx` com a **primeira linha como cabeçalho** e selecione a coluna que deseja usar para separar os dados.
Os arquivos separados manterão a **formatação visual original** (cores, bordas, estilos etc.).
""")

uploaded_file = st.file_uploader("📁 Envie seu arquivo .xlsx", type=["xlsx"])

if uploaded_file:
    try:
        # Lê uma prévia com pandas apenas para mostrar ao usuário
        df_preview = pd.read_excel(uploaded_file, nrows=5)
        df_preview = df_preview.dropna(axis=1, how="all")
        df_preview = df_preview.loc[:, ~df_preview.columns.str.contains('^Unnamed')]
        
        coluna_separadora = st.selectbox(
            "Selecione a coluna para separar os arquivos:",
            options=df_preview.columns,
            index=0
        )

        st.success(f"Arquivo carregado com sucesso! Coluna selecionada: **{coluna_separadora}**")
        st.write("Visualização da planilha (5 primeiras linhas):")
        st.write(df_preview)

        if st.button("🚀 Separar e baixar arquivos com formatação"):
            input_excel = BytesIO(uploaded_file.read())
            wb_original = load_workbook(input_excel)
            ws_original = wb_original.active

            # Cabeçalho
            colunas = [cell.value for cell in ws_original[1]]

            # Índice da coluna escolhida
            idx_coluna_sep = colunas.index(coluna_separadora) + 1

            # Agrupar linhas por valor da coluna
            dados_por_valor = {}
            for row in ws_original.iter_rows(min_row=2, values_only=False):
                valor = row[idx_coluna_sep - 1].value
                if valor:
                    chave = str(valor).strip().lower()
                    if chave not in dados_por_valor:
                        dados_por_valor[chave] = []
                    dados_por_valor[chave].append(row)

            # Criar zip com arquivos formatados
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for chave, linhas in dados_por_valor.items():
                    wb_novo = Workbook()
                    ws_novo = wb_novo.active

                    # Copiar cabeçalho
                    for col_idx, cell in enumerate(ws_original[1], start=1):
                        if cell.value is None:
                            continue  # pula colunas Unnamed/vazias
                        novo_cell = ws_novo.cell(row=1, column=col_idx, value=cell.value)
                        if cell.has_style:
                            novo_cell._style = cell._style

                    # Copiar dados com estilo
                    for row_idx, row in enumerate(linhas, start=2):
                        for col_idx, cell in enumerate(row, start=1):
                            header = ws_original.cell(row=1, column=col_idx).value
                            if header is None:
                                continue  # ignora colunas vazias/Unnamed
                            novo_cell = ws_novo.cell(row=row_idx, column=col_idx, value=cell.value)
                            if cell.has_style:
                                novo_cell._style = cell._style

                    # Criar arquivo em memória
                    nome_arquivo = f"{chave}.xlsx".replace("/", "_").replace("\\", "_").replace(":", "-")
                    excel_bytes = BytesIO()
                    wb_novo.save(excel_bytes)
                    excel_bytes.seek(0)
                    zip_file.writestr(nome_arquivo, excel_bytes.read())

            zip_buffer.seek(0)
            st.download_button(
                label="📦 Baixar arquivos separados com formatação (.zip)",
                data=zip_buffer,
                file_name="planilhas_separadas.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
