import streamlit as st
import pandas as pd
import io

# Configuração
st.set_page_config(page_title="Auditoria Interna NF - Produto", layout="wide")

st.title("📊 Auditoria Interna NF - Produto")

# --- UPLOAD ---
col1, col2 = st.columns(2)
with col1:
    file_nf_prod = st.file_uploader("1. Relatório de NF's", type=['xlsx'])
    file_forn = st.file_uploader("2. Relatório de Credores", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato", type=['xlsx', 'csv'])
    file_titulo = st.file_uploader("6. Relatório Título", type=['xlsx'])

# --- FUNÇÕES ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def limpar_cod(v):
    if pd.isna(v): return ""
    return str(v).split('.')[0].strip().lstrip('0')

def extrair_nf_produto(v):
    if pd.isna(v): return ""
    v = str(v).split('/')[0]
    v = "".join(filter(str.isdigit, v))
    return v.lstrip('0')

def extrair_nf_painel(v):
    if pd.isna(v): return ""
    v = str(v)
    if '/' in v:
        v = v.split('/')[-1]
    v = "".join(filter(str.isdigit, v))
    return v.lstrip('0')

# 🔥 TITULO
def tratar_titulo(file):
    df_raw = pd.read_excel(file, header=None)

    start_idx = None
    for i in range(len(df_raw)):
        if str(df_raw.iloc[i, 0]).strip() == "Item":
            start_idx = i
            break

    if start_idx is None:
        return pd.DataFrame()

    df = df_raw.iloc[start_idx+1:].copy()
    df.columns = df_raw.iloc[start_idx]

    df = df[['Credor', 'Documento', 'Titulo', 'CT/OC', 'Emis.NF', 'Valor líquido']].copy()

    df['Credor'] = df['Credor'].astype(str).str.strip()
    df['Documento'] = df['Documento'].astype(str).str.strip()
    df['CT/OC'] = df['CT/OC'].astype(str).str.strip()
    df['Emis.NF'] = pd.to_datetime(df['Emis.NF'], errors='coerce')
    df['Valor líquido'] = pd.to_numeric(df['Valor líquido'], errors='coerce').fillna(0)

    df['NF'] = df['Documento'].str.replace(r'\D', '', regex=True).str.lstrip('0')

    # chave do boleto
    df['chave_boleto'] = (
        df['Credor'].str.upper() + "_" +
        df['CT/OC'] + "_" +
        df['Emis.NF'].astype(str)
    )

    soma = df.groupby('chave_boleto')['Valor líquido'].sum().reset_index()
    soma = soma.rename(columns={'Valor líquido': 'Valor Boleto'})

    df = df.merge(soma, on='chave_boleto', how='left')

    return df

# --- EXECUÇÃO ---
if st.button("🚀 Iniciar Auditoria"):

    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato, file_titulo]):

        # --- LEITURA ---
        df_nf = pd.read_excel(file_nf_prod)
        df_forn = pd.read_excel(file_forn)
        df_painel = pd.read_excel(file_painel)
        df_relacao = pd.read_excel(file_relacao)
        df_titulo = tratar_titulo(file_titulo)

        # --- CNPJ ---
        df_forn['CNPJ/CPF'] = df_forn['CNPJ/CPF'].apply(limpar_cnpj)
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.upper()

        # --- TITULO + CNPJ ---
        df_titulo['Credor_UP'] = df_titulo['Credor'].str.upper()

        titulo_cnpj = pd.merge(
            df_titulo,
            df_forn[['Credor_UP', 'CNPJ/CPF']],
            on='Credor_UP',
            how='left'
        )

        titulo_cnpj['CNPJ/CPF'] = titulo_cnpj['CNPJ/CPF'].apply(limpar_cnpj)

        # --- PAINEL ---
        df_painel['nf_ref_limpa'] = df_painel['N° da Nota fiscal'].apply(extrair_nf_painel)

        titulo_full = pd.merge(
            titulo_cnpj,
            df_painel[['nf_ref_limpa', 'N° do Pedido']],
            left_on='NF',
            right_on='nf_ref_limpa',
            how='left'
        )

        # --- FINAL ---
        auditoria_titulo = titulo_full.rename(columns={
            'N° do Pedido': 'Nº do pedido',
            'NF': 'NF',
            'CNPJCPF': 'CNPJ',
            'Credor': 'Credor',
            'Emis.NF': 'Data emissão',
            'Valor líquido': 'Valor',
            'Valor Boleto': 'Valor boleto'
        })[
            ['Nº do pedido', 'NF', 'CNPJ', 'Credor', 'Data emissão', 'Valor', 'Valor boleto']
        ]

        # --- EXPORTAÇÃO ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:

            # mantém seu padrão original (se quiser pode plugar aqui suas abas antigas)
            auditoria_titulo.to_excel(writer, sheet_name='4. TITULO', index=False)

        st.success("✅ Auditoria gerada com sucesso!")
        st.download_button(
            "📥 Baixar Auditoria",
            output.getvalue(),
            "AUDITORIA_TITULO.xlsx"
        )
