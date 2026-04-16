import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF - Produtos", layout="wide")

st.title("📊 Auditoria Interna NF (Notas de Produtos)")
st.markdown("Carregue o arquivo **bruto** de notas de produtos e os demais relatórios.")

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
    """Remove pontuação e garante que o CNPJ tenha 14 dígitos (ou CPF 11)"""
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    if not num: return ""
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def limpar_cod(v):
    if pd.isna(v): return ""
    return str(v).split('.')[0].strip().lstrip('0')

def extrair_nf(v):
    if pd.isna(v) or v == "": return ""
    # Pega apenas o número antes de qualquer barra ou traço
    return "".join(filter(str.isdigit, str(v).split('/')[0])).strip()

def estruturar_notas_produtos_interno(file):
    df_bruto = pd.read_excel(file, header=None)
    registros = []
    cnpj_dest = None
    colunas_id = None
    processando = False

    for i, row in df_bruto.iterrows():
        val_a = str(row[0]).strip() if pd.notna(row[0]) else ""
        if "CNPJ do destinatário:" in val_a:
            # Limpamos o CNPJ do destinatário aqui também
            cnpj_dest = limpar_cnpj(row[3])
            processando = False
            continue
        if val_a == "Emitente":
            colunas_id = [str(c).strip() for c in row.values]
            processando = True
            continue
        if processando and val_a != "" and val_a != "nan":
            registros.append([cnpj_dest] + list(row.values))

    df = pd.DataFrame(registros, columns=['CNPJ Destinatário'] + colunas_id)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed|^nan|None', case=False, na=False)]
    return df.dropna(subset=['Emitente'])

def transformar_credor_limpo(df_bruto):
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
        
        # 1. Estruturação Interna (Notas de Produtos)
        df_nf = estruturar_notas_produtos_interno(file_nf_prod)
        
        # 2. Carga dos outros arquivos
        df_forn = transformar_credor_limpo(pd.read_excel(file_forn, header=None))
        df_painel = pd.read_excel(file_painel)
        df_relacao = pd.read_excel(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None)

        # --- PADRONIZAÇÃO DE CHAVES (CRUCIAL) ---
        df_forn['CNPJCPF'] = df_forn['CNPJCPF'].apply(limpar_cnpj)
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj)
        
        # Criamos a chave de busca (CNPJ + Número NF)
        df_nf['nf_limpa'] = df_nf['Núm/Série'].apply(extrair_nf)
        df_nf['chave_unica'] = df_nf['CNPJ emitente'] + "_" + df_nf['nf_limpa']

        # --- CRUZAMENTO PAINEL ---
        df_painel['nf_ref_limpa'] = df_painel['N° da Nota fiscal'].apply(extrair_nf)
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.strip().str.upper()
        
        # Trazemos o CNPJ para o Painel para poder criar a chave de comparação
        painel_com_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJCPF'].apply(limpar_cnpj) + "_" + painel_com_cnpj['nf_ref_limpa']
        
        chaves_lancadas = set(painel_com_cnpj[painel_com_cnpj['nf_ref_limpa'] != ""]['chave_p'].unique())
        
        resumo_painel = pd.merge(df_nf, painel_com_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), left_on='chave_unica', right_on='chave_p', how='left')
        
        resumo_painel['Status'] = resumo_painel.apply(lambda r: "✅ NF Lançada" if r['chave_unica'] in chaves_lancadas else "❌ Sem Histórico", axis=1)

        # --- CRUZAMENTO PEDIDOS ---
        df_relacao['Cód. fornecedor'] = df_relacao['Cód. fornecedor'].apply(limpar_cod)
        rel_com_cnpj = pd.merge(df_relacao, df_forn[['Cód. Fornecedor', 'CNPJCPF']], left_on='Cód. fornecedor', right_on='Cód. Fornecedor', how='left')
        
        if 'CNPJCPF' in rel_com_cnpj.columns and 'Nº do pedido' in rel_com_cnpj.columns:
            rel_com_cnpj['CNPJCPF'] = rel_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
            peds_agrupados = rel_com_cnpj.dropna(subset=['Nº do pedido']).groupby('CNPJCPF')['Nº do pedido'].apply(
                lambda x: ", ".join(sorted(set(x.astype(str).unique())))
            ).reset_index()
        else:
            peds_agrupados = pd.DataFrame(columns=['CNPJCPF', 'Nº do pedido'])

        resumo_pedidos = pd.merge(resumo_painel, peds_agrupados, left_on='CNPJ emitente', right_on='CNPJCPF', how='left')

        # --- CRUZAMENTO CONTRATOS ---
        registros_ct = []
        item_atual = {'Contrato': None, 'CNPJ': None}
        for i in range(len(df_bruto_ct)):
            l = df_bruto_ct.iloc[i]
            col_a = str(l[0]).strip() if pd.notna(l[0]) else ""
            if col_a == "Contrato":
                item_atual['Contrato'] = str(l[3]).strip()
            elif col_a == "CNPJ" and item_atual['Contrato']:
                item_atual['CNPJ'] = limpar_cnpj(l[3]) # Garante CNPJ limpo no contrato
                registros_ct.append(item_atual.copy())
        
        if registros_ct:
            cts_agrupados = pd.DataFrame(registros_ct).groupby('CNPJ')['Contrato'].apply(
                lambda x: ", ".join(sorted(set(x.astype(str).unique())))
            ).reset_index()
        else:
            cts_agrupados = pd.DataFrame(columns=['CNPJ', 'Contrato'])

        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on='CNPJ emitente', right_on='CNPJ', how='left')

        # --- COLUNAS FINAIS ---
        cols_base = ['Núm/Série', 'CNPJ emitente', 'Emitente', 'Emissão', 'Valor']
        cols_extra = ['CNPJ Destinatário', 'Destinatário']
        
        aba1 = resumo_painel[cols_base + ['N° da Nota fiscal', 'Status'] + cols_extra]
        aba2 = resumo_pedidos[cols_base + ['N° da Nota fiscal', 'Nº do pedido', 'Status'] + cols_extra]
        aba3 = resumo_contratos[cols_base + ['N° da Nota fiscal', 'Nº do pedido', 'Contrato', 'Status'] + cols_extra]

        # Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Auditoria processada! Verifique se os dados agora foram encontrados.")
        st.download_button("📥 Baixar Relatório Final", output.getvalue(), "AUDITORIA_PRODUTOS.xlsx")
    else:
        st.error("Carregue todos os arquivos.")
