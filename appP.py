import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF - Produtos", layout="wide")

st.title("📊 Auditoria Interna NF (Notas de Produtos)")
st.markdown("Carregue o arquivo **bruto** de notas de produtos e os demais relatórios para auditoria.")

# --- UPLOAD DOS FICHEIROS ---
col1, col2 = st.columns(2)
with col1:
    file_nf_prod = st.file_uploader("1. Relatório de Notas de Produtos (Arquivo BRUTO)", type=['xlsx'])
    file_forn = st.file_uploader("2. Relatório de Credores", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato", type=['xlsx', 'csv'])

# --- FUNÇÕES DE LIMPEZA E ESTRUTURAÇÃO ---

def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def limpar_cod(v):
    if pd.isna(v): return ""
    return str(v).split('.')[0].strip().lstrip('0')

def extrair_nf(v):
    if pd.isna(v) or v == "": return ""
    # Para produtos, remove o que vem após a barra (ex: 454/1 -> 454)
    return str(v).split('/')[0].strip()

def estruturar_notas_produtos_interno(file):
    """
    Lógica solicitada: Lê o arquivo bruto, captura CNPJ do Destinatário 
    e limpa colunas 'nan' ou 'Unnamed'.
    """
    df_bruto = pd.read_excel(file, header=None)
    registros = []
    cnpj_atual = None
    colunas_identificadas = None
    processando_tabela = False

    for i, row in df_bruto.iterrows():
        valor_col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
        
        # GATILHO 1: Captura o CNPJ do destinatário na coluna D
        if "CNPJ do destinatário:" in valor_col_a:
            cnpj_atual = str(row[3]).strip() if pd.notna(row[3]) else ""
            processando_tabela = False
            continue
        
        # GATILHO 2: Identifica o cabeçalho da tabela
        if valor_col_a == "Emitente":
            colunas_identificadas = [str(c).strip() for c in row.values]
            processando_tabela = True
            continue
        
        # PROCESSAMENTO: Captura os dados enquanto a tabela durar
        if processando_tabela and valor_col_a != "" and valor_col_a != "nan":
            linha_dados = [cnpj_atual] + list(row.values)
            registros.append(linha_dados)

    # Monta o DataFrame inicial
    header_final = ['CNPJ Destinatário'] + colunas_identificadas
    df_final = pd.DataFrame(registros, columns=header_final)

    # Limpeza de colunas fantasmas (nan, Unnamed)
    df_final = df_final.loc[:, ~df_final.columns.str.contains('^Unnamed|^nan|None', case=False, na=False)]
    df_final = df_final.dropna(axis=1, how='all')
    
    # Filtro final de segurança
    return df_final.dropna(subset=['Emitente'])

def transformar_credor_limpo(df_bruto):
    # (Mantendo sua lógica de credores para garantir o vínculo de CNPJ)
    for i in range(min(15, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            df_header = df_header.loc[:, df_header.columns.notna() & (df_header.columns != 'nan')]
            def split_safe(val):
                s = str(val).strip()
                return (s.split(" - ")[0], " - ".join(s.split(" - ")[1:])) if " - " in s else ("", s)
            res_split = df_header['Credor'].apply(split_safe)
            df_header['Cód. Fornecedor'] = res_split.apply(lambda x: x[0])
            df_header['Fornecedor'] = res_split.apply(lambda x: x[1])
            return df_header.rename(columns={'CNPJ/CPF': 'CNPJCPF'})
    return df_bruto

# --- PROCESSAMENTO PRINCIPAL ---
if st.button("🚀 Iniciar Auditoria"):
    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato]):
        
        # ETAPA 1: Tratamento interno das notas brutas
        df_nf = estruturar_notas_produtos_interno(file_nf_prod)
        
        # ETAPA 2: Carga dos demais arquivos
        df_forn = transformar_credor_limpo(pd.read_excel(file_forn, header=None))
        df_painel = pd.read_excel(file_painel)
        df_relacao = pd.read_excel(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None)

        # Padronização de chaves
        df_forn['CNPJCPF'] = df_forn['CNPJCPF'].apply(limpar_cnpj)
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf['Núm/Série'].apply(extrair_nf)
        df_nf['chave_unica'] = df_nf['CNPJ emitente'] + "_" + df_nf['nf_limpa']

        # --- CRUZAMENTO PAINEL ---
        df_painel['nf_ref_limpa'] = df_painel['N° da Nota fiscal'].apply(extrair_nf)
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.strip().str.upper()
        
        painel_com_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJCPF'] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        chaves_lancadas = set(painel_com_cnpj[painel_com_cnpj['nf_ref_limpa'] != ""]['chave_p'].unique())
        
        resumo_painel = pd.merge(df_nf, painel_com_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), left_on='chave_unica', right_on='chave_p', how='left')
        resumo_painel['Status'] = resumo_painel.apply(lambda r: "✅ NF Lançada" if r['chave_unica'] in chaves_lancadas else "❌ Sem Histórico", axis=1)

        # --- CRUZAMENTO PEDIDOS (VERSÃO CORRIGIDA) ---
if 'CNPJCPF' in rel_com_cnpj.columns and 'Nº do pedido' in rel_com_cnpj.columns:
    # Removemos nulos antes de agrupar para evitar o TypeError no join
    rel_com_cnpj_clean = rel_com_cnpj.dropna(subset=['CNPJCPF', 'Nº do pedido'])
    
    peds_agrupados = rel_com_cnpj_clean.groupby('CNPJCPF')['Nº do pedido'].apply(
        lambda x: ", ".join(sorted(set(x.astype(str).unique())))
    ).reset_index()
else:
    peds_agrupados = pd.DataFrame(columns=['CNPJCPF', 'Nº do pedido'])

        # --- CRUZAMENTO CONTRATOS ---
        registros_ct = []
        item_atual = {'Contrato': None, 'CNPJ': None}
        for i in range(len(df_bruto_ct)):
            l = df_bruto_ct.iloc[i]
            col_a = str(l[0]).strip() if pd.notna(l[0]) else ""
            if col_a == "Contrato": item_atual['Contrato'] = str(l[3]).strip()
            elif col_a == "CNPJ" and item_atual['Contrato']:
                item_atual['CNPJ'] = limpar_cnpj(l[3])
                registros_ct.append(item_atual.copy())
        
        cts_agrupados = pd.DataFrame(registros_ct).groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index() if registros_ct else pd.DataFrame(columns=['CNPJ', 'Contrato'])
        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on='CNPJ emitente', right_on='CNPJ', how='left')

        # Seleção de Colunas Final (Sequência solicitada)
        cols_base = ['Núm/Série', 'CNPJ emitente', 'Emitente', 'Emissão', 'Valor']
        cols_extra = ['CNPJ Destinatário', 'Destinatário']
        
        # Geração das abas
        aba1 = resumo_painel[cols_base + ['N° da Nota fiscal', 'Status'] + cols_extra]
        aba2 = resumo_pedidos[cols_base + ['N° da Nota fiscal', 'Nº do pedido', 'Status'] + cols_extra]
        aba3 = resumo_contratos[cols_base + ['N° da Nota fiscal', 'Nº do pedido', 'Contrato', 'Status'] + cols_extra]

        # Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Auditoria processada com sucesso!")
        st.download_button("📥 Baixar Relatório Final", output.getvalue(), "AUDITORIA_PRODUTOS.xlsx")
    else:
        st.error("Por favor, carregue todos os arquivos.")
