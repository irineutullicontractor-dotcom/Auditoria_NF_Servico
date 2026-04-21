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
    # Padroniza nomes das colunas (remove espaços e garante strings)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def transformar_credor_limpo(file):
    """Localiza a linha 'Credor' e 'CNPJ/CPF' na planilha Credores"""
    df_bruto = pd.read_excel(file, header=None)
    for i in range(min(25, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = row_values
            df_header = df_header.loc[:, df_header.columns.notna() & (df_header.columns != 'nan')]
            return df_header
    return pd.DataFrame()

# --- INTERFACE ---
st.set_page_config(page_title="Auditoria Título", layout="wide")
st.title("📑 Auditoria Interna - Títulos")
st.markdown("""
### Instruções de uso:
1. Carregue o relatório do **Painel** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
2. Carregue o relatório de **Pedidos** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
3. Carregue o relatório de **Títulos** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
4. Carregue o relatório de **Credores**.
""")

col1, col2 = st.columns(2)
with col1:
    file_painel = st.file_uploader("1. Relatório Painel - Home / Suprimentos / Compras / Painel de Compras (Novo).", type=['xlsx'])
    file_pedidos = st.file_uploader("2. Relatório Pedidos - Home / Suprimentos / Compras / Relatórios / Pedidos de compra / Relação de Pedidos de Compra (Novo).", type=['xlsx'])
with col2:
    file_titulo = st.file_uploader("3. Relatório Titulo - Home / Financeiro / Contas a Pagar / Relatórios / Títulos por Data.", type=['xlsx'])
    file_credor = st.file_uploader("4. Relatório de Credores - Home / Mais Opções / Apoio / Relatórios / Pessoas / Credores.", type=['xlsx'])

if st.button("🚀 Gerar Auditoria Atualizada"):
    if all([file_painel, file_pedidos, file_titulo, file_credor]):
        try:
            # 1. Processar Credores
            df_c = transformar_credor_limpo(file_credor)
            if df_c.empty:
                st.error("Erro: Colunas 'Credor' e 'CNPJ/CPF' não encontradas na planilha de Credores.")
                st.stop()
            df_c['CNPJ_LIMPO'] = df_c['CNPJ/CPF'].apply(limpar_cnpj)
            
            # 2. Processar Título
            df_t = estruturar_titulo_limpo(file_titulo)
            if df_t.empty:
                st.error("Erro: Linha de início 'Item' não encontrada na planilha Título.")
                st.stop()

            # 3. Cruzamento para CNPJ
            df_t = pd.merge(df_t, df_c[['Credor', 'CNPJ_LIMPO']].drop_duplicates('Credor'), on='Credor', how='left')

            # 4. Lógica de Soma do Boleto
            # Identifica as colunas de valor e título (com ou sem acento)
            col_valor = 'Valor líquido' if 'Valor líquido' in df_t.columns else df_t.columns[-1]
            col_titulo_orig = 'Titulo' if 'Titulo' in df_t.columns else ('Título' if 'Título' in df_t.columns else None)
            
            if col_titulo_orig:
                df_t['Valor boleto'] = df_t.groupby(['CT/OC', 'Credor', 'Emis.NF'])[col_valor].transform('sum')

                # 5. Montagem do arquivo final com a ordem das colunas solicitada
                resumo = pd.DataFrame()
                resumo['Nº do pedido'] = df_t['CT/OC']
                resumo['NF'] = df_t['Documento']
                resumo['CNPJ'] = df_t['CNPJ_LIMPO']
                resumo['Credor'] = df_t['Credor']
                resumo['Data emissão'] = df_t['Emis.NF']
                resumo['Valor'] = df_t[col_valor]
                resumo['Titulo'] = df_t[col_titulo_orig] # Coluna adicionada aqui
                resumo['Valor boleto'] = df_t['Valor boleto']

                # 6. Exportação
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    resumo.to_excel(writer, sheet_name='auditoria_titulo', index=False)
                
                st.success("✅ Tudo pronto! Relatório de Auditoria gerado com sucesso.")
                st.download_button("📥 Baixar Auditoria", output.getvalue(), "AUDITORIA_TITULO.xlsx")
            else:
                st.error("Coluna 'Titulo' não encontrada na planilha de origem.")

        except Exception as e:
            st.error(f"Erro inesperado: {e}")
    else:
        st.info("Por favor, faça o upload de todos os arquivos para prosseguir.")
