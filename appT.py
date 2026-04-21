import streamlit as st
import pandas as pd
import io

# --- FUNÇÕES DE APOIO ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def estruturar_titulo_limpo(file):
    """
    Limpa a planilha Título: localiza a linha que começa com 'Item' 
    e extrai os dados úteis.
    """
    df_bruto = pd.read_excel(file, header=None)
    
    # Localizar onde os dados começam (Coluna A == 'Item')
    inicio_dados = None
    for i, row in df_bruto.iterrows():
        if str(row[0]).strip().lower() == "item":
            inicio_dados = i
            break
    
    if inicio_dados is None:
        return pd.DataFrame()

    # Define a linha encontrada como cabeçalho
    df = pd.read_excel(file, skiprows=inicio_dados)
    
    # Selecionar apenas colunas necessárias para evitar ruído
    cols_necessarias = ['Credor', 'Documento', 'Titulo', 'CT/OC', 'Emis.NF', 'Valor líquido']
    # Filtro de segurança: manter apenas se as colunas existirem
    df = df[[c for c in cols_necessarias if c in df.columns]]
    
    # Limpeza básica: remove linhas totalmente vazias ou que repetem o cabeçalho
    df = df.dropna(subset=['Credor', 'Documento'])
    return df

# --- PROCESSAMENTO PRINCIPAL ---
st.title("📑 Auditoria Título - Integração de Planilhas")

# Upload dos 4 arquivos
file_painel = st.file_uploader("1. Carregue o Painel", type=['xlsx'])
file_pedidos = st.file_uploader("2. Carregue os Pedidos", type=['xlsx'])
file_titulo = st.file_uploader("3. Carregue o Titulo", type=['xlsx'])
file_credor = st.file_uploader("4. Carregue o Credor", type=['xlsx'])

if st.button("Gerar Auditoria Titulo"):
    if all([file_painel, file_pedidos, file_titulo, file_credor]):
        
        # 1. Leitura e Limpeza Inicial
        df_p = pd.read_excel(file_painel)
        df_ped = pd.read_excel(file_pedidos)
        df_t = estruturar_titulo_limpo(file_titulo)
        df_c = pd.read_excel(file_credor)

        # 2. Padronização de Chaves
        # No Credor: CNPJ e Nome
        df_c['CNPJ_LIMPO'] = df_c['CNPJ/CPF'].apply(limpar_cnpj)
        
        # No Titulo: Precisamos do CNPJ que está na planilha Credor
        # Fazemos um merge para trazer o CNPJ para a planilha Título usando o nome do Credor
        df_t = pd.merge(df_t, df_c[['Credor', 'CNPJ_LIMPO']], on='Credor', how='left')

        # 3. Lógica do "Valor Boleto" (Agrupamento)
        # Regra: CT/OC igual + Credor igual + Data Emissão igual = Mesmo Boleto
        # Vamos criar uma coluna de soma agrupada
        df_t['Valor boleto'] = df_t.groupby(['CT/OC', 'Credor', 'Emis.NF'])['Valor líquido'].transform('sum')

        # 4. Integração com Pedidos e Painel para pegar o Nº do Pedido e NF
        # Nota: Título usa 'Documento' como NF e 'CT/OC' como Pedido.
        # Vamos preparar o DataFrame Final
        
        final_df = pd.DataFrame()
        final_df['Nº do pedido'] = df_t['CT/OC']
        final_df['NF'] = df_t['Documento']
        final_df['CNPJ'] = df_t['CNPJ_LIMPO']
        final_df['Credor'] = df_t['Credor']
        final_df['Data emissão'] = df_t['Emis.NF']
        final_df['Valor'] = df_t['Valor líquido']
        final_df['Valor boleto'] = df_t['Valor boleto']

        # 5. Exportação
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='auditoria_titulo', index=False)
        
        st.success("Arquivo auditoria_titulo gerado!")
        st.download_button("📥 Baixar Planilha Integrada", output.getvalue(), "auditoria_titulo.xlsx")
    else:
        st.error("Por favor, carregue todos os 4 arquivos.")
