# app.py

import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO

# ========== CONFIGURAÇÃO DA PÁGINA ==========
st.set_page_config(
    page_title="Soma de Entradas e Saídas (CFOP)",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Soma de Valores de Entrada e Saída por CFOP")
st.write(
    """
    Faça o upload de um ou mais arquivos de planilha (Excel ou CSV). O app irá:
    1. Ler cada planilha a partir da linha 18 (pular cabeçalhos ou linhas iniciais).
    2. Tratar a coluna **CFOP** como texto.
    3. Tratar a coluna **Valor NF** corretamente (se já for numérico, manter; 
       se for string no estilo brasileiro, remover pontos de milhar e trocar vírgula por ponto).
    4. Classificar cada registro em **Entrada** (CFOP começando com 1, 2 ou 3) ou **Saída** (CFOP começando com 5, 6 ou 7).
    5. Somar os valores de cada categoria.
    6. Exibir um resumo detalhado por sheet de cada arquivo e um resumo agregado por nome de arquivo.
    7. Exibir os totais gerais de todas as planilhas combinadas.
    """
)

# ========== FUNÇÃO PARA PROCESSAR CADA ARQUIVO ==========
def processar_arquivo(file_buffer: BytesIO, filename: str) -> pd.DataFrame:
    """
    Lê o arquivo (Excel ou CSV) e devolve um DataFrame com a soma de
    valores de entrada e saída para cada planilha/sheet desse arquivo.
    
    Retorna um DataFrame com colunas:
    ['arquivo', 'sheet', 'total_entrada', 'total_saida']
    """
    resultados = []
    ext = Path(filename).suffix.lower()

    # --- PARA EXCEL: ler todas as sheets ---
    if ext in [".xlsx", ".xls"]:
        sheets_dict = pd.read_excel(
            file_buffer,
            sheet_name=None,
            skiprows=17,            # pular até a linha 17 (linha 18 vira cabeçalho)
            dtype={"CFOP": str},    # força CFOP como texto
            engine="openpyxl"
        )
        for sheet_name, df in sheets_dict.items():
            resultado = processar_dataframe(df, filename, sheet_name)
            resultados.append(resultado)

    # --- PARA CSV: lê apenas como um único DataFrame ---
    elif ext == ".csv":
        df = pd.read_csv(
            file_buffer,
            skiprows=17,
            dtype={"CFOP": str},   # força CFOP como texto
            sep=None,              # autodetectar delimitador
            engine="python"
        )
        resultado = processar_dataframe(df, filename, "CSV")
        resultados.append(resultado)

    else:
        st.error(f"⚠️ Formato não suportado: {ext}")
        return pd.DataFrame()

    # Concatena resultados de todas as sheets lidas do arquivo
    if resultados:
        return pd.concat(resultados, ignore_index=True)
    else:
        return pd.DataFrame()


def processar_dataframe(df: pd.DataFrame, arquivo: str, sheet: str) -> pd.DataFrame:
    """
    Recebe o DataFrame lido (já a partir da linha 18) e devolve
    um DataFrame com soma de entradas e saídas para aquela sheet.
    """
    # Renomear colunas para remover espaços acidentais
    df = df.rename(columns=lambda x: str(x).strip())

    # Verifica se as colunas mínimas estão presentes
    colunas_necessarias = {"CFOP", "Valor NF"}
    faltantes = colunas_necessarias - set(df.columns)
    if faltantes:
        st.warning(f"No arquivo **{arquivo}**, sheet **{sheet}** faltam colunas: {faltantes}")
        return pd.DataFrame({
            "arquivo": [arquivo],
            "sheet": [sheet],
            "total_entrada": [0.0],
            "total_saida": [0.0]
        })

    # Garante que CFOP seja string e retira espaços
    df["CFOP"] = df["CFOP"].astype(str).str.strip()

    # ————— TRATAMENTO CORRETO DE "Valor NF" —————
    if pd.api.types.is_numeric_dtype(df["Valor NF"]):
        # Já é numérico, basta preencher NaN
        df["Valor NF"] = df["Valor NF"].fillna(0.0)
    else:
        # Se veio como string (ex.: "1.234,56"), remover pontos de milhar, trocar vírgula por ponto
        df["Valor NF"] = (
            df["Valor NF"].astype(str)
            .str.replace(r"\.", "", regex=True)     # remove pontos (milhares)
            .str.replace(",", ".", regex=False)     # vírgula → ponto decimal
            .str.replace(r"[^\d\.-]", "", regex=True)  # remove qualquer outro caractere que não seja dígito, ponto ou hífen
        )
        df["Valor NF"] = pd.to_numeric(df["Valor NF"], errors="coerce").fillna(0.0)
    # ——— FIM DO TRATAMENTO ———

    # Máscaras para entrada (CFOP 1xx, 2xx, 3xx) e saída (CFOP 5xx, 6xx, 7xx)
    mask_entrada = df["CFOP"].str.startswith(("1", "2", "3"))
    mask_saida   = df["CFOP"].str.startswith(("5", "6", "7"))

    total_entrada = df.loc[mask_entrada, "Valor NF"].sum()
    total_saida   = df.loc[mask_saida,   "Valor NF"].sum()

    return pd.DataFrame({
        "arquivo": [arquivo],
        "sheet":   [sheet],
        "total_entrada": [total_entrada],
        "total_saida":   [total_saida]
    })


# ========== UI: UPLOADER DE ARQUIVOS ==========
st.sidebar.header("📂 Upload de Arquivos")
arquivos = st.sidebar.file_uploader(
    "Selecione uma ou mais planilhas (Excel ou CSV)",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

if not arquivos:
    st.warning("Faça o upload de ao menos um arquivo para iniciar a análise.")
    st.stop()

# Botão para processar
if st.sidebar.button("▶️ Processar arquivos"):
    todos_resultados = []
    barra = st.progress(0)

    for idx, uploaded_file in enumerate(arquivos):
        # Lê em BytesIO para permitir múltiplas leituras (Excel precisa disso)
        file_bytes = uploaded_file.read()
        file_buffer = BytesIO(file_bytes)
        filename = uploaded_file.name

        df_res = processar_arquivo(file_buffer, filename)
        if not df_res.empty:
            todos_resultados.append(df_res)

        barra.progress((idx + 1) / len(arquivos))

    # Concatena resultados de todos os arquivos
    if todos_resultados:
        df_final = pd.concat(todos_resultados, ignore_index=True)
    else:
        df_final = pd.DataFrame(
            columns=["arquivo", "sheet", "total_entrada", "total_saida"]
        )

    # ===================================================
    # 1) Exibição: resumo detalhado por arquivo e sheet
    # ===================================================
    st.markdown("## 📑 Resumo Detalhado por Arquivo e Sheet")
    st.dataframe(df_final.style.format({
        "total_entrada": "R$ {:,.2f}",
        "total_saida":   "R$ {:,.2f}"
    }), height=350)

    # ===================================================
    # 2) Exibição: resumo agregado por NOME DE ARQUIVO
    # ===================================================
    if not df_final.empty:
        df_por_arquivo = (
            df_final
            .groupby("arquivo", as_index=False)
            .agg({
                "total_entrada": "sum",
                "total_saida":   "sum"
            })
            .rename(columns={
                "total_entrada": "soma_entrada_no_arquivo",
                "total_saida":   "soma_saida_no_arquivo"
            })
        )

        st.markdown("## 📂 Resumo Agregado por Nome de Arquivo")
        st.dataframe(df_por_arquivo.style.format({
            "soma_entrada_no_arquivo": "R$ {:,.2f}",
            "soma_saida_no_arquivo":   "R$ {:,.2f}"
        }), height=250)
    else:
        st.info("Nenhum resultado para agrupar por arquivo.")

    # ===================================================
    # 3) Totais gerais de todas as planilhas combinadas
    # ===================================================
    total_entrada_geral = df_final["total_entrada"].sum()
    total_saida_geral   = df_final["total_saida"].sum()

    st.markdown("---")
    st.markdown("## 📌 Totais Gerais (todos os arquivos)")
    st.markdown(f"- **Entrada (CFOP 1xx, 2xx, 3xx):** R$ {total_entrada_geral:,.2f}")
    st.markdown(f"- **Saída (CFOP 5xx, 6xx, 7xx):**   R$ {total_saida_geral:,.2f}")

    # ===================================================
    # 4) Download CSV com todos os resultados (sheet-level)
    # ===================================================
    csv_export = df_final.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Baixar resultados consolidados (.csv)",
        data=csv_export,
        file_name="resumo_cfop_entradas_saidas.csv",
        mime="text/csv"
    )
