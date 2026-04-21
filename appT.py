import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auditoria Ultra Blindada", layout="wide")
st.title("📊 Auditoria Título - Ultra Blindada")

# --- UPLOAD ---
file_forn = st.file_uploader("Credores", type=['xlsx', 'csv'])
file_painel = st.file_uploader("Painel", type=['xlsx', 'csv'])
file_titulo = st.file_uploader("Título", type=['xlsx'])

# --- UTIL ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14)

def extrair_nf(v):
    if pd.isna(v): return ""
    return "".join(filter(str.isdigit, str(v))).lstrip("0")

def normalizar(df):
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

# 🔥 DETECTOR DE CABEÇALHO POR SCORE
def detectar_header(df_raw, palavras_chave):
    melhor_idx = None
    melhor_score = 0

    for i in range(min(30, len(df_raw))):
        row = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
        score = sum(any(p in cell for p in palavras_chave) for cell in row)

        if score > melhor_score:
            melhor_score = score
            melhor_idx = i

    return melhor_idx

# 🔥 ENCONTRAR COLUNA FLEXÍVEL
def get_col(df, nomes):
    for col in df.columns:
        for n in nomes:
            if n in col:
                return col
    return None

# --- CREDOR ---
def tratar_credor(file):
    df_raw = pd.read_excel(file, header=None)

    idx = detectar_header(df_raw, ["CREDOR", "CNPJ", "CPF"])

    if idx is None:
        st.error("❌ Não encontrou cabeçalho credor")
        st.stop()

    df = df_raw.iloc[idx+1:].copy()
    df.columns = df_raw.iloc[idx]
    df = normalizar(df)

    col_credor = get_col(df, ["CREDOR"])
    col_cnpj = get_col(df, ["CNPJ", "CPF"])

    df = df[[col_credor, col_cnpj]].copy()
    df.columns = ["CREDOR", "CNPJ"]

    df["CREDOR"] = df["CREDOR"].astype(str).str.upper().str.strip()
    df["CNPJ"] = df["CNPJ"].apply(limpar_cnpj)

    return df

# --- TITULO ---
def tratar_titulo(file):
    df_raw = pd.read_excel(file, header=None)

    idx = detectar_header(df_raw, ["ITEM", "DOCUMENTO", "VALOR", "CREDOR"])

    if idx is None:
        st.error("❌ Não encontrou cabeçalho título")
        st.stop()

    df = df_raw.iloc[idx+1:].copy()
    df.columns = df_raw.iloc[idx]
    df = normalizar(df)

    col_credor = get_col(df, ["CREDOR"])
    col_doc = get_col(df, ["DOCUMENTO"])
    col_ct = get_col(df, ["CT", "OC"])
    col_data = get_col(df, ["EMIS", "DATA"])
    col_valor = get_col(df, ["VALOR"])

    df = df[[col_credor, col_doc, col_ct, col_data, col_valor]].copy()
    df.columns = ["CREDOR", "DOCUMENTO", "CTOC", "DATA", "VALOR"]

    df["CREDOR"] = df["CREDOR"].astype(str).str.upper().str.strip()
    df["NF"] = df["DOCUMENTO"].apply(extrair_nf)
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0)

    # 🔥 BOLETO
    df["CHAVE"] = df["CREDOR"] + "_" + df["CTOC"].astype(str) + "_" + df["DATA"].astype(str)

    soma = df.groupby("CHAVE")["VALOR"].sum().reset_index()
    soma.columns = ["CHAVE", "VALOR_BOLETO"]

    df = df.merge(soma, on="CHAVE", how="left")

    return df

# --- PAINEL ---
def tratar_painel(file):
    df = pd.read_excel(file)
    df = normalizar(df)

    col_nf = get_col(df, ["NOTA"])
    col_pedido = get_col(df, ["PEDIDO"])

    df = df[[col_nf, col_pedido]].copy()
    df.columns = ["NF", "PEDIDO"]

    df["NF"] = df["NF"].apply(extrair_nf)

    return df

# --- EXECUÇÃO ---
if st.button("🚀 Rodar Auditoria Ultra"):

    if all([file_forn, file_painel, file_titulo]):

        df_forn = tratar_credor(file_forn)
        df_titulo = tratar_titulo(file_titulo)
        df_painel = tratar_painel(file_painel)

        # JOIN
        df = pd.merge(df_titulo, df_forn, on="CREDOR", how="left")
        df = pd.merge(df, df_painel, on="NF", how="left")

        auditoria = df.rename(columns={
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

        # EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            auditoria.to_excel(writer, sheet_name="TITULO", index=False)

        st.success("✅ Ultra blindado concluído")
        st.download_button("📥 Baixar", output.getvalue(), "auditoria_ultra.xlsx")
