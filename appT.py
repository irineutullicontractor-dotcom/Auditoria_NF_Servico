import streamlit as st
import pandas as pd
import io

# --- FUNÇÕES DE LIMPEZA ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def estruturar_titulo_limpo(file):
    """Localiza a linha 'Item' e limpa a planilha Titulo"""
    df_bruto = pd.read_excel(file, header=None)
    inicio_dados = None
    for i, row in df_bruto.iterrows():
        if str(row[0]).strip().lower() == "item":
            inicio_dados = i
            break
    
    if inicio_dados is None:
        return pd.DataFrame()

    df = pd.read_excel(file, skiprows=inicio_dados)
    # Remove colunas totalmente vazias e garante nomes de colunas sem espaços
    df.columns = [str(c).strip() for c in df.columns]
    return df

def transformar_credor_limpo(file):
    """Localiza a linha 'Credor' e 'CNPJ/CPF' na planilha Credores"""
    df_bruto = pd.read_excel(file, header=None)
    for i in range(min(20, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = row_values
            # Filtra colunas nulas ou 'nan'
            df_header = df_header.loc[:, df_header.columns.notna() & (df_header.columns != 'nan')]
            return df_header
    return pd.DataFrame()

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Auditoria Título", layout="wide")
st.title("📑 Auditoria Título - Integração de Planilhas")

file_painel = st.file_uploader("1. Carregue o Painel", type=['xlsx'])
file_pedidos = st.file_uploader("2. Carregue os Pedidos", type=['xlsx'])
file_titulo = st.file_uploader("3. Carregue o Titulo", type=['xlsx'])
file_credor = st.file_uploader("4. Carregue o Credor", type=['xlsx'])

if st.button("🚀 Gerar Auditoria"):
    if all([file_painel, file_pedidos, file_titulo, file_credor]):
        try:
            # 1. Processamento da planilha Credor (onde dava o erro)
            df_c = transformar_credor_limpo(file_credor)
            if df_c.empty:
                st.error("Não foi possível encontrar as colunas 'Credor' e 'CNPJ/CPF' na planilha de Credores.")
                st.stop()
            
            df_c['CNPJ_LIMPO'] = df_c['CNPJ/CPF'].apply(limpar_cnpj)
            
            # 2. Processamento da planilha Título
            df_t = estruturar_titulo_limpo(file_titulo)
            if df_t.empty:
                st.error("Não foi possível encontrar a linha de início ('Item') na planilha Título.")
                st.stop()

            # 3. Cruzamento para trazer o CNPJ para o Título
            # (Usamos o nome do Credor para vincular o CNPJ)
            df_t = pd.merge(df_t, df_c[['Credor', 'CNPJ_LIMPO']].drop_duplicates('Credor'), on='Credor', how='left')

            # 4. Lógica de Soma do Boleto (Agrupamento)
            # Agrupa por CT/OC (Pedido), Credor e Data de Emissão
            col_valor = 'Valor líquido' if 'Valor líquido' in df_t.columns else df_t.columns[-1]
            df_t['Valor boleto'] = df_t.groupby(['CT/OC', 'Credor', 'Emis.NF'])[col_valor].transform('sum')

            # 5. Montagem do arquivo final
            # Mapeamento conforme solicitado
            resumo = pd.DataFrame()
            resumo['Nº do pedido'] = df_t['CT/OC']
            resumo['NF'] = df_t['Documento']
            resumo['CNPJ'] = df_t['CNPJ_LIMPO']
            resumo['Credor'] = df_t['Credor']
            resumo['Data emissão'] = df_t['Emis.NF']
            resumo['Valor'] = df_t[col_valor]
            resumo['Valor boleto'] = df_t['Valor boleto']

            # 6. Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resumo.to_excel(writer, sheet_name='auditoria_titulo', index=False)
            
            st.success("✅ Integração concluída com sucesso!")
            st.download_button("📥 Baixar auditoria_titulo.xlsx", output.getvalue(), "auditoria_titulo.xlsx")

        except Exception as e:
            st.error(f"Erro durante o processamento: {e}")
    else:
        st.warning("Aguardando o upload de todos os arquivos.")
