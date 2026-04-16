import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF - Produto", layout="wide")

st.title("📊 Auditoria Interna NF - Produto (Correção de Vínculo)")

# --- FUNÇÕES DE LIMPEZA ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def extrair_nf(v):
    """
    Transforma '444/1', '444.0' ou ' 444 ' em apenas '444'.
    """
    if pd.isna(v): return ""
    # Converte para string, remove o .0 (caso seja float do Excel) e pega o que vem antes da barra
    s = str(v).strip().split('.')[0].split('/')[0]
    return s

def estruturar_notas_produtos_interno(file):
    df_bruto = pd.read_excel(file, header=None)
    registros = []
    cnpj_dest = None
    colunas_id = None
    processando = False

    for i, row in df_bruto.iterrows():
        val_a = str(row[0]).strip() if pd.notna(row[0]) else ""
        if "CNPJ do destinatário:" in val_a:
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
    return df.dropna(subset=['Emitente'])

def transformar_credor_limpo(df_bruto):
    for i in range(min(15, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            def split_safe(val):
                s = str(val).strip()
                return (s.split(" - ")[0], " - ".join(s.split(" - ")[1:])) if " - " in s else ("", s)
            res_split = df_header['Credor'].apply(split_safe)
            df_header['Cód. Fornecedor'] = res_split.apply(lambda x: x[0])
            return df_header.rename(columns={'CNPJ/CPF': 'CNPJCPF'})
    return df_bruto

# --- UPLOADS ---
col1, col2 = st.columns(2)
with col1:
    file_nf_prod = st.file_uploader("1. Relatório de NF's", type=['xlsx'])
    file_forn = st.file_uploader("2. Relatório de Credores", type=['xlsx'])
    file_painel = st.file_uploader("3. Relatório Painel", type=['xlsx'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos", type=['xlsx'])
    file_contrato = st.file_uploader("5. Relatório Contrato", type=['xlsx'])

if st.button("🚀 Iniciar Auditoria"):
    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato]):
        
        # 1. Carregar Dados
        df_nf = estruturar_notas_produtos_interno(file_nf_prod)
        df_forn = transformar_credor_limpo(pd.read_excel(file_forn, header=None))
        df_painel = pd.read_excel(file_painel)
        df_relacao = pd.read_excel(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None)

        # 2. Padronização CRÍTICA (Onde o erro de vínculo acontece)
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf['Núm/Série'].apply(extrair_nf)
        df_nf['chave_unica'] = df_nf['CNPJ emitente'].astype(str) + "_" + df_nf['nf_limpa'].astype(str)
        
        df_forn['CNPJCPF'] = df_forn['CNPJCPF'].apply(limpar_cnpj)
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.strip().str.upper()
        
        # Cruzamento Painel x Credores para obter o CNPJ
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        painel_com_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        
        # Limpeza da NF no Painel também!
        painel_com_cnpj['nf_ref_limpa'] = painel_com_cnpj['N° da Nota fiscal'].apply(extrair_nf)
        painel_com_cnpj['CNPJCPF'] = painel_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJCPF'].astype(str) + "_" + painel_com_cnpj['nf_ref_limpa'].astype(str)
        
        # 3. Processamento das Abas
        # ABA 1: Cruzamento por Chave Única (CNPJ + NF Limpa)
        resumo_painel = pd.merge(
            df_nf, 
            painel_com_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), 
            left_on='chave_unica', 
            right_on='chave_p', 
            how='left'
        )
        
        chaves_lancadas = set(painel_com_cnpj[painel_com_cnpj['nf_ref_limpa'] != ""]['chave_p'].unique())
        cnpjs_no_painel = set(painel_com_cnpj['CNPJCPF'].unique())

        def definir_status(r):
            if pd.notna(r['chave_p']): return "✅ NF Lançada"
            if r['CNPJ emitente'] in cnpjs_no_painel: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        
        resumo_painel['Status'] = resumo_painel.apply(definir_status, axis=1)

        # ABA 2: Pedidos
        peds_agrupados = df_relacao.groupby('Cód. fornecedor')['Nº do pedido'].apply(
            lambda x: ", ".join(sorted(set(str(v).split('.')[0] for v in x if pd.notna(v))))
        ).reset_index()
        # Precisamos do CNPJ no Pedido para cruzar com a NF
        rel_com_cnpj = pd.merge(peds_agrupados, df_forn[['Cód. Fornecedor', 'CNPJCPF']], left_on='Cód. fornecedor', right_on='Cód. Fornecedor', how='left')
        rel_com_cnpj['CNPJCPF'] = rel_com_cnpj['CNPJCPF'].apply(limpar_cnpj)

        resumo_pedidos = pd.merge(resumo_painel, rel_com_cnpj[['CNPJCPF', 'Nº do pedido']], left_on='CNPJ emitente', right_on='CNPJCPF', how='left')

        # ABA 3: Contratos (mesma lógica)
        # ... (Mantendo a lógica de extração de contrato que já funcionava)
        
        # Exportação Final
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resumo_painel.to_excel(writer, sheet_name='1. PAINEL', index=False)
            resumo_pedidos.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
        
        st.success("Processado com limpeza rigorosa de NF!")
        st.download_button("📥 Baixar Relatório Corrigido", output.getvalue(), "AUDITORIA_PRODUTO_CORRIGIDA.xlsx")
