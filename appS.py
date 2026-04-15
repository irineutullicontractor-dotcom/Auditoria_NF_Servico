import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Master 3 em 1", layout="wide")

st.title("📊 Auditoria Master: Painel, Pedidos & Contratos")
st.markdown("""
### Fluxo de Conferência:
O sistema processará as Notas Fiscais em três níveis:
1. **Aba 1 (Painel):** Match exato de NF com o Painel.
2. **Aba 2 (Pedidos):** Notas pendentes cruzadas com a Relação de Pedidos (Oficina).
3. **Aba 3 (Contratos):** Notas pendentes cruzadas com o Relatório de Contratos.
""")

# --- UPLOAD DOS 5 FICHEIROS ---
col1, col2 = st.columns(2)
with col1:
    file_nf = st.file_uploader("1. Relatório de NFs", type=['xlsx', 'csv'])
    file_forn = st.file_uploader("2. Cadastro de Fornecedores", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Relação/Oficina (Pedidos)", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato (BRUTO)", type=['xlsx', 'csv'])

def carregar(file, header=0):
    if file is None: return None
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    return pd.read_excel(file, header=header)

if st.button("🚀 Iniciar Auditoria Completa"):
    if not all([file_nf, file_forn, file_painel, file_relacao, file_contrato]):
        st.error("Por favor, carregue os 5 arquivos para continuar.")
    else:
        # Carregamento
        df_nf = carregar(file_nf)
        df_forn = carregar(file_forn)
        df_painel = carregar(file_painel)
        df_relacao = carregar(file_relacao)
        df_bruto_contrato = carregar(file_contrato, header=None)

        # --- MAPEAMENTO INTELIGENTE E FLEXÍVEL ---
        def encontrar_coluna(df, opcoes):
            for opt in opcoes:
                if opt in df.columns: return opt
            return None

        NF_CNPJ = encontrar_coluna(df_nf, ['CNPJ Prestador (CNPJ)', 'Prestador (CNPJ)', 'Prestador (CNPJ / CPF)', 'CNPJ'])
        NF_NUMERO = encontrar_coluna(df_nf, ['Número NFS-e (nNFSe)', 'Número (nNFSe)', 'nNFSe'])
        NF_FORN = encontrar_coluna(df_nf, ['Nome Prestador (xNome)', 'Prestador (xNome)', 'Razão Social Prestador'])
        NF_DATA = encontrar_coluna(df_nf, ['Data/Hora Emissão DPS (dhEmi)', 'Data da Emissão (dhEmi)', 'dhEmi'])
        NF_VALOR = encontrar_coluna(df_nf, ['Valor do Serviço (vServ) (vServ)', 'Valor Serviço (vServ)', 'vServ'])

        if not NF_CNPJ:
            st.error(f"❌ Coluna de CNPJ não encontrada nas NFs. Colunas lidas: {list(df_nf.columns)}")
            st.stop()

        # Padrões para os outros arquivos
        PED_FORN_PAINEL, PED_NUM_PAINEL, PED_NF_REF = 'Fornecedor', 'N° do Pedido', 'N° da Nota fiscal'
        PED_FORN_REL, PED_NUM_REL = 'Cód. fornecedor', 'Nº do pedido'
        FORN_COD, FORN_CNPJ, FORN_CRED = 'Cód. Fornecedor', 'CNPJCPF', 'Credor'

        # --- PADRONIZAÇÕES ---
        def limpar_cnpj(v):
            num = "".join(filter(str.isdigit, str(v)))
            return num.zfill(14) if len(num) > 11 else num.zfill(11)

        def limpar_cod(v):
            return str(v).split('.')[0].strip().lstrip('0')

        def extrair_nf(v):
            if pd.isna(v) or v == "": return ""
            return "".join(filter(str.isdigit, str(v).split('/')[-1])).strip()

        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf[NF_NUMERO].astype(str).str.strip()
        df_forn[FORN_CNPJ] = df_forn[FORN_CNPJ].apply(limpar_cnpj)
        df_forn[FORN_COD] = df_forn[FORN_COD].apply(limpar_cod)
        df_forn[FORN_CRED] = df_forn[FORN_CRED].str.strip().str.upper()

        # --- PROCESSAMENTO DO CONTRATO BRUTO ---
        registros_ct = []
        item_atual = {}
        for i, row in df_bruto_contrato.iterrows():
            col_a, col_c, col_d = [str(row[idx]).strip() if pd.notna(row[idx]) else "" for idx in [0, 2, 3]]
            if col_c == "Contrato":
                if item_atual: registros_ct.append(item_atual)
                item_atual = {'Contrato': None, 'CNPJ': None}
            if item_atual:
                if col_a == "Contrato": item_atual['Contrato'] = col_d
                elif col_a == "CNPJ": item_atual['CNPJ'] = col_d
        if item_atual: registros_ct.append(item_atual)
        df_contrato_limpo = pd.DataFrame(registros_ct).dropna(how='all')
        df_contrato_limpo['CNPJ'] = df_contrato_limpo['CNPJ'].apply(limpar_cnpj)

        # =================================================================
        # ABA 1: PAINEL
        # =================================================================
        df_painel[PED_FORN_PAINEL] = df_painel[PED_FORN_PAINEL].str.strip().str.upper()
        df_painel['nf_extraida'] = df_painel[PED_NF_REF].apply(extrair_nf)
        painel_com_cnpj = pd.merge(df_painel, df_forn[[FORN_CRED, FORN_CNPJ]], left_on=PED_FORN_PAINEL, right_on=FORN_CRED, how='left')
        
        df_nf['chave'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']
        painel_com_cnpj['chave'] = painel_com_cnpj[FORN_CNPJ] + "_" + painel_com_cnpj['nf_extraida']
        
        match_painel = pd.merge(df_nf, painel_com_cnpj, on='chave', how='inner')
        match_painel['Status'] = "✅ NF Lançada"
        
        nfs_restantes = df_nf[~df_nf['chave'].isin(match_painel['chave'])]
        sugestoes_painel = pd.merge(nfs_restantes, painel_com_cnpj, left_on=NF_CNPJ, right_on=FORN_CNPJ, how='left')
        sugestoes_painel['Status'] = sugestoes_painel[PED_NUM_PAINEL].apply(lambda x: "⚠️ Pedido Encontrado" if pd.notna(x) else "❌ Sem Pedido")
        
        resumo_painel = pd.concat([match_painel, sugestoes_painel], ignore_index=True)
        resumo_painel = resumo_painel[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, PED_NUM_PAINEL, PED_NF_REF, NF_VALOR, 'Status']].drop_duplicates()

        # =================================================================
        # ABA 2: PEDIDOS (OFICINA)
        # =================================================================
        df_relacao[PED_FORN_REL] = df_relacao[PED_FORN_REL].apply(limpar_cod)
        rel_com_cnpj = pd.merge(df_relacao, df_forn[[FORN_COD, FORN_CNPJ]], left_on=PED_FORN_REL, right_on=FORN_COD, how='left')
        peds_agrupados = rel_com_cnpj.groupby(FORN_CNPJ)[PED_NUM_REL].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()
        
        resumo_pedidos = pd.merge(df_nf, peds_agrupados, left_on=NF_CNPJ, right_on=FORN_CNPJ, how='left')
        nfs_resolvidas_painel = match_painel['nf_limpa'].unique()
        
        def status_pedidos(r):
            if r['nf_limpa'] in nfs_resolvidas_painel: return "✅ Resolvido no Painel"
            return "⚠️ Pendente (Oficina)" if pd.notna(r[PED_NUM_REL]) else "❌ Sem Pedido Oficina"
        
        resumo_pedidos['Status'] = resumo_pedidos.apply(status_pedidos, axis=1)
        resumo_pedidos = resumo_pedidos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, PED_NUM_REL, NF_VALOR, 'Status']]

        # =================================================================
        # ABA 3: CONTRATOS
        # =================================================================
        cts_agrupados = df_contrato_limpo.groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()
        resumo_contratos = pd.merge(df_nf, cts_agrupados, left_on=NF_CNPJ, right_on='CNPJ', how='left')
        
        def status_contratos(r):
            # Só marca como resolvido se estiver no match exato do PAINEL
            if r['nf_limpa'] in nfs_resolvidas_painel: return "✅ Resolvido no Painel"
            return "⚠️ Contrato Encontrado" if pd.notna(r['Contrato']) else "❌ Sem Contrato"
            
        resumo_contratos['Status'] = resumo_contratos.apply(status_contratos, axis=1)
        resumo_contratos = resumo_contratos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, 'Contrato', NF_VALOR, 'Status']]

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resumo_painel.to_excel(writer, sheet_name='1. PAINEL', index=False)
            resumo_pedidos.to_excel(writer, sheet_name='2. PEDIDOS (OFICINA)', index=False)
            resumo_contratos.to_excel(writer, sheet_name='3. CONTRATOS', index=False)
        
        st.success("Auditoria Master concluída com sucesso!")
        st.download_button(label="📥 Baixar Relatório Master 3-em-1", data=output.getvalue(), file_name="AUDITORIA_MASTER_COMPLETA.xlsx", mime="application/vnd.ms-excel")
