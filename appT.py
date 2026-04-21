import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auditoria Título Blindada", layout="wide")
st.title("📊 Auditoria Título (Versão Blindada)")

# --- UPLOAD ---
file_forn = st.file_uploader("Credores", type=['xlsx', 'csv'])
file_painel = st.file_uploader("Painel", type=['xlsx', 'csv'])
file_titulo = st.file_uploader("Título", type=['xlsx'])

# --- FUNÇÕES BASE ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14)

def extrair_nf(v):
    if pd.isna(v): return ""
    v = "".join(filter(str.isdigit, str(v)))
    return v.lstrip('0')

def normalizar_colunas(df):
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def encontrar_coluna(df, possiveis_nomes):
    for col in df.columns:
        for nome in possiveis_nomes:
            if nome in col:
                return col
    return None

# --- 🔥 CREDOR BLINDADO ---
def tratar_credor(file):
    df_raw = pd.read_excel(file, header=None)

    # encontrar linha de cabeçalho
    header_idx = None
    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.upper()
        if any("CREDOR" in x for x in row) and any("CNPJ" in x for x in row):
            header_idx = i
            break

    if header_idx is None:
        st.error("❌ Não foi possível identificar cabeçalho de Credores")
        st.stop()

    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = df_raw.iloc[header_idx]
    df = normalizar_colunas(df)

    col_credor = encontrar_coluna(df, ["CREDOR"])
    col_cnpj = encontrar_coluna(df, ["CNPJ", "CPF"])

    df = df[[col_credor, col_cnpj]].copy()
    df.columns = ["CREDOR", "CNPJ"]

    df["CREDOR"] = df["CREDOR"].astype(str).str.strip().str.upper()
    df["CNPJ"] = df["CNPJ"].apply(limpar_cnpj)

    return df

# --- 🔥 TITULO BLINDADO ---
def tratar_titulo(file):
    df_raw = pd.read_excel(file, header=None)

    start_idx = None
    for i in range(len(df_raw)):
        if str(df_raw.iloc[i, 0]).strip().upper() == "ITEM":
            start_idx = i
            break

    if start_idx is None:
        st.error("❌ Não encontrou início da planilha Título (ITEM)")
        st.stop()

    df = df_raw.iloc[start_idx+1:].copy()
    df.columns = df_raw.iloc[start_idx]
    df = normalizar_colunas(df)

    col_credor = encontrar_coluna(df, ["CREDOR"])
    col_doc = encontrar_coluna(df, ["DOCUMENTO"])
    col_ct = encontrar_coluna(df, ["CT/OC", "OC"])
    col_data = encontrar_coluna(df, ["EMIS"])
    col_valor = encontrar_coluna(df, ["VALOR"])

    df = df[[col_credor, col_doc, col_ct, col_data, col_valor]].copy()
    df.columns = ["CREDOR", "DOCUMENTO", "CTOC", "DATA", "VALOR"]

    df["CREDOR"] = df["CREDOR"].astype(str).str.upper().str.strip()
    df["NF"] = df["DOCUMENTO"].apply(extrair_nf)
    df["DATA"] = pd.to_datetime(df["DATA"], errors='coerce')
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors='coerce').fillna(0)

    # 🔥 REGRA DO BOLETO
    df["CHAVE"] = df["CREDOR"] + "_" + df["CTOC"].astype(str) + "_" + df["DATA"].astype(str)

    soma = df.groupby("CHAVE")["VALOR"].sum().reset_index()
    soma.columns = ["CHAVE", "VALOR_BOLETO"]

    df = df.merge(soma, on="CHAVE", how="left")

    return df

# --- 🔥 PAINEL BLINDADO ---
def tratar_painel(file):
    df = pd.read_excel(file)
    df = normalizar_colunas(df)

    col_nf = encontrar_coluna(df, ["NOTA"])
    col_pedido = encontrar_coluna(df, ["PEDIDO"])

    df = df[[col_nf, col_pedido]].copy()
    df.columns = ["NF", "PEDIDO"]

    df["NF"] = df["NF"].apply(extrair_nf)

    return df

# --- EXECUÇÃO ---
if st.button("🚀 Rodar Auditoria"):

    if all([file_forn, file_painel, file_titulo]):

        df_forn = tratar_credor(file_forn)
        df_titulo = tratar_titulo(file_titulo)
        df_painel = tratar_painel(file_painel)

        # --- JOIN CNPJ ---
        df_final = pd.merge(
            df_titulo,
            df_forn,
            on="CREDOR",
            how="left"
        )

        # --- JOIN PEDIDO ---
        df_final = pd.merge(
            df_final,
            df_painel,
            on="NF",
            how="left"
        )

        # --- FINAL ---
        auditoria = df_final.rename(columns={
            "PEDIDO": "Nº do pedido",
            "NF": "NF",
            "CNPJ": "CNPJ",
            "CREDOR": "Credor",
            "DATA": "Data emissão",
            "VALOR": "Valor",
            "VALOR_BOLETO": "Valor boleto"
        })[
            ["Nº do pedido", "NF", "CNPJ", "Credor", "Data emissão", "Valor", "Valor boleto"]
        ]

        # --- EXPORT ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            auditoria.to_excel(writer, sheet_name='TITULO', index=False)

        st.success("✅ Auditoria pronta (versão blindada)")
        st.download_button("📥 Baixar", output.getvalue(), "auditoria_titulo.xlsx")
