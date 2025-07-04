import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from openpyxl import load_workbook, Workbook
from copy import copy

st.set_page_config(page_title="Separador de Planilhas com Formatação", layout="centered")
st.title("📊 Separador de Planilha com ou sem Formatação")

st.markdown("""
Envie um arquivo Excel `.xlsx` com a **primeira linha como cabeçalho** e selecione a coluna para separar os dados.

- Os arquivos separados manterão a **formatação visual original** (cores, bordas, estilos).
- Fórmulas serão convertidas em **valores fixos** — evitando nomes estranhos de arquivo e linhas vazias.
- Se algo falhar, há um segundo botão para baixar **sem formatação**.
""")

uploaded_file = st.file_uploader("📁 Envie seu arquivo .xlsx", type=["xlsx"])

if uploaded_file:
    try:
        # Carrega o binário UMA única vez
        raw_bytes = BytesIO(uploaded_file.read())
        raw_bytes.seek(0)

        # Pré‑visualização (pandas sempre lê valores, não fórmulas)
        df_preview = pd.read_excel(raw_bytes, nrows=5)
        df_preview = (df_preview
                      .dropna(axis=1, how="all")
                      .loc[:, ~df_preview.columns.str.contains('^Unnamed')])

        coluna_separadora = st.selectbox(
            "Selecione a coluna para separar os arquivos:",
            options=df_preview.columns,
            index=0
        )

        st.success(f"Arquivo carregado. Coluna selecionada: **{coluna_separadora}**")
        st.write("Visualização (5 primeiras linhas):")
        st.write(df_preview)

        # ------------------------------------------------------------------ #
        # 1) Botão principal – manter formatação & converter fórmulas em valores
        # ------------------------------------------------------------------ #
        if st.button("✨ Separar e baixar COM formatação"):
            try:
                # Precisamos de duas cópias independentes do buffer
                buf_fmt  = BytesIO(raw_bytes.getvalue()); buf_fmt.seek(0)
                buf_vals = BytesIO(raw_bytes.getvalue()); buf_vals.seek(0)

                wb_fmt  = load_workbook(buf_fmt,  data_only=False)  # estilos + fórmulas
                wb_vals = load_workbook(buf_vals, data_only=True)   # apenas valores
                ws_fmt  = wb_fmt.active
                ws_vals = wb_vals.active

                # Índice da coluna separadora
                headers = [c.value for c in ws_fmt[1]]
                idx_sep = headers.index(coluna_separadora) + 1  # 1‑based

                # Agrupa linhas por valor (já sem fórmulas)
                grupos = {}
                for r_fmt, r_val in zip(ws_fmt.iter_rows(min_row=2, values_only=False),
                                        ws_vals.iter_rows(min_row=2, values_only=False)):
                    chave = r_val[idx_sep-1].value
                    if chave is None or str(chave).strip() == "":
                        continue
                    chave_norm = str(chave).strip().lower()
                    grupos.setdefault(chave_norm, []).append((r_fmt, r_val, chave))

                # Cria ZIP
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for chave_norm, linhas in grupos.items():
                        # Usa o valor original (primeira ocorrência) para o nome
                        nome_arquivo = str(linhas[0][2]).strip()
                        nome_arquivo = (nome_arquivo
                                        .replace("/", "_")
                                        .replace("\\", "_")
                                        .replace(":", "-"))
                        if not nome_arquivo:
                            nome_arquivo = "vazio"

                        wb_new = Workbook()
                        ws_new = wb_new.active

                        # --- Cabeçalhos ---
                        for c_idx, (c_fmt, c_val) in enumerate(zip(ws_fmt[1], ws_vals[1]), start=1):
                            if c_fmt.value is None:
                                continue
                            tgt = ws_new.cell(row=1, column=c_idx, value=c_fmt.value)
                            if c_fmt.has_style:
                                tgt.font       = copy(c_fmt.font)
                                tgt.fill       = copy(c_fmt.fill)
                                tgt.border     = copy(c_fmt.border)
                                tgt.alignment  = copy(c_fmt.alignment)
                                tgt.number_format = copy(c_fmt.number_format)

                        # --- Linhas de dados ---
                        for r_idx, (row_fmt, row_val, _) in enumerate(linhas, start=2):
                            for c_idx, (cell_fmt, cell_val) in enumerate(zip(row_fmt, row_val), start=1):
                                # pula colunas vazias no cabeçalho
                                if ws_fmt.cell(row=1, column=c_idx).value is None:
                                    continue
                                tgt = ws_new.cell(row=r_idx, column=c_idx, value=cell_val.value)
                                if cell_fmt.has_style:
                                    tgt.font       = copy(cell_fmt.font)
                                    tgt.fill       = copy(cell_fmt.fill)
                                    tgt.border     = copy(cell_fmt.border)
                                    tgt.alignment  = copy(cell_fmt.alignment)
                                    tgt.number_format = copy(cell_fmt.number_format)

                        # Salva no ZIP
                        bytes_out = BytesIO()
                        wb_new.save(bytes_out)
                        bytes_out.seek(0)
                        zip_file.writestr(f"{nome_arquivo}.xlsx", bytes_out.read())

                zip_buffer.seek(0)
                st.download_button(
                    "📥 Baixar arquivos separados COM formatação (.zip)",
                    data=zip_buffer,
                    file_name="planilhas_formatadas.zip",
                    mime="application/zip"
                )

            except Exception as e:
                st.error(f"Erro ao manter formatação/valores: {e}")
                st.info("Tente o botão alternativo abaixo para baixar sem formatação.")

        # ------------------------------------------------------------------ #
        # 2) Botão alternativo – sem formatação (mais rápido, à prova de erro)
        # ------------------------------------------------------------------ #
        if st.button("📁 Separar e baixar SEM formatação"):
            try:
                raw_bytes.seek(0)
                df_full = pd.read_excel(raw_bytes).dropna(axis=1, how="all")
                df_full = df_full.loc[:, ~df_full.columns.str.contains('^Unnamed')]

                df_full['__key'] = (df_full[coluna_separadora]
                                    .astype(str)
                                    .str.strip()
                                    .str.lower())

                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for chave_norm, g in df_full.groupby('__key'):
                        if chave_norm == "" or chave_norm.lower() == "nan":
                            continue
                        valor_original = g[coluna_separadora].iloc[0]
                        nome = str(valor_original).strip()
                        nome = nome.replace("/", "_").replace("\\", "_").replace(":", "-")
                        if not nome:
                            nome = "vazio"

                        bytes_out = BytesIO()
                        g.drop(columns='__key').to_excel(bytes_out, index=False)
                        bytes_out.seek(0)
                        zip_file.writestr(f"{nome}.xlsx", bytes_out.read())

                zip_buffer.seek(0)
                st.download_button(
                    "📥 Baixar arquivos separados SEM formatação (.zip)",
                    data=zip_buffer,
                    file_name="planilhas_sem_formatacao.zip",
                    mime="application/zip"
                )
            except Exception as e:
                st.error(f"Erro ao separar sem formatação: {e}")

    except Exception as e:
        st.error(f"Não foi possível ler o arquivo: {e}")
