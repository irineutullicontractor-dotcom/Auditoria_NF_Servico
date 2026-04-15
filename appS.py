import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Master: Fluxo de Pendências", layout="wide")

st.title("📊 Auditoria Master: Fluxo Painel -> Pedidos -> Contratos")
st.markdown("""
### Como funciona o fluxo:
- **Aba 1 (Painel):** Exibe todas as notas e tenta o match exato.
- **Aba 2 (Pedidos):** Mostra apenas o que **não** foi lançado no Painel, tentando encontrar pedidos da oficina.
- **Aba 3 (Contratos):** O filtro final. Mostra apenas o que **não** está no Painel nem na Oficina, tentando o vínculo por Contrato.
""")

# --- UPLOAD DOS 5 FICHEIROS ---
col1, col2 = st.columns(2)
with col1:
    file_nf = st.file_uploader("1. Relatório de NFs", type=['xlsx', 'csv'])
    file_forn = st.file_uploader("2. Cadastro de Fornecedores", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos (Oficina)", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato (BRUTO)", type=['xlsx', 'csv'])

def carregar(file, header=0):
    if file is None: return None
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    return pd.read_excel(file, header=header)

if st.button("🚀 Gerar Auditoria em Funil"):
    if not all([file_nf, file_forn, file_painel, file_relacao, file_contrato]):
        st.error("Por favor, carregue os 5 arquivos.")
    else:
        # Carregamento
        df_nf = carregar(file_nf)
        df_forn = carregar(file_forn)
        df_painel = carregar(file_painel)
        df_relacao = carregar(file_relacao)
        df_bruto_contrato = carregar(file_contrato, header=None)

        # --- MAPEAMENTO FLEXÍVEL ---
        def encontrar_coluna(df, opcoes):
            df.columns = df.columns.str.strip()
            for opt in opcoes:
                if opt in df.columns: return opt
            return None

        NF_CNPJ = encontrar_coluna(df_nf, ['CNPJ Prestador (CNPJ)', 'Prestador (CNPJ)', 'Prestador (CNPJ / CPF)', 'CNPJ'])
        NF_NUMERO = encontrar_coluna(df_nf, ['Número NFS-e (nNFSe)', 'Número (nNFSe)', 'nNFSe'])
        NF_FORN = encontrar_coluna(df_nf, ['Nome Prestador (xNome)', 'Prestador (xNome)', 'Razão Social Prestador'])
        NF_DATA = encontrar_coluna(df_nf, ['Data/Hora Emissão DPS (dhEmi)', 'Data da Emissão (dhEmi)', 'dhEmi'])
        NF_VALOR = encontrar_coluna(df_nf, ['Valor do Serviço (vServ) (vServ)', 'Valor Serviço (vServ)', 'vServ'])
        
        PED_FORN_REL = encontrar_coluna(df_relacao, ['Cód. fornecedor', 'Cód. Fornecedor', 'Fornecedor', 'Código'])
        PED_NUM_REL = encontrar_coluna(df_relacao, ['Nº do pedido', 'N° do Pedido', 'Pedido', 'Número'])
        
        PED_FORN_PAINEL, PED_NUM_PAINEL, PED_NF_REF = 'Fornecedor', 'N° do Pedido', 'N° da Nota fiscal'
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

        # Limpeza Contratos
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
        df_ct_limpo = pd.DataFrame(registros_ct).dropna(how='all')
        df_ct_limpo['CNPJ'] = df_ct_limpo['CNPJ'].apply(limpar_cnpj)

        # =================================================================
        # ABA 1: PAINEL (BASE COMPLETA)
        # =================================================================
        df_painel[PED_FORN_PAINEL] = df_painel[PED_FORN_PAINEL].str.strip().str.upper()
        df_painel['nf_extraida'] = df_painel[PED_NF_REF].apply(extrair_nf)
        painel_com_cnpj = pd.merge(df_painel, df_forn[[FORN_CRED, FORN_CNPJ]], left_on=PED_FORN_PAINEL, right_on=FORN_CRED, how='left')
        
        df_nf['chave'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']
        painel_com_cnpj['chave'] = painel_com_cnpj[FORN_CNPJ] + "_" + painel_com_cnpj['nf_extraida']
        
        match_painel = pd.merge(df_nf, painel_com_cnpj, on='chave', how='inner')
        nfs_resolvidas_painel = match_painel['nf_limpa'].unique()
        
        # Gerar a aba 1
        resumo_painel = pd.merge(df_nf, painel_com_cnpj, on='chave', how='left')
        resumo_painel['Status'] = resumo_painel.apply(lambda r: "✅ NF Lançada" if r['nf_limpa'] in nfs_resolvidas_painel else ("⚠️ Pedido Encontrado" if pd.notna(r[PED_NUM_PAINEL]) else "❌ Sem Pedido"), axis=1)
        resumo_painel = resumo_painel[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, PED_NUM_PAINEL, NF_VALOR, 'Status']].drop_duplicates()

        # =================================================================
        # ABA 2: PEDIDOS (APENAS PENDENTES DO PAINEL)
        # =================================================================
        df_pendentes_painel = df_nf[~df_nf['nf_limpa'].isin(nfs_resolvidas_painel)].copy()
        
        df_relacao[PED_FORN_REL] = df_relacao[PED_FORN_REL].apply(limpar_cod)
        rel_com_cnpj = pd.merge(df_relacao, df_forn[[FORN_COD, FORN_CNPJ]], left_on=PED_FORN_REL, right_on=FORN_COD, how='left')
        peds_agrupados = rel_com_cnpj.groupby(FORN_CNPJ)[PED_NUM_REL].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()
        
        resumo_pedidos = pd.merge(df_pendentes_painel, peds_agrupados, left_on=NF_CNPJ, right_on=FORN_CNPJ, how='left')
        
        # Notas resolvidas na Oficina
        nfs_com_pedido_oficina = resumo_pedidos[resumo_pedidos[PED_NUM_REL].notna()]['nf_limpa'].unique()
        
        resumo_pedidos['Status'] = resumo_pedidos[PED_NUM_REL].apply(lambda x: "⚠️ Pendente (Oficina)" if pd.notna(x) else "❌ Sem Pedido Oficina")
        resumo_pedidos = resumo_pedidos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, PED_NUM_REL, NF_VALOR, 'Status']]

        # =================================================================
        # ABA 3: CONTRATOS (APENAS PENDENTES DE TUDO)
        # =================================================================
        # Pega o que não está nem no Painel nem na Oficina
        df_pendentes_total = df_pendentes_painel[~df_pendentes_painel['nf_limpa'].isin(nfs_com_pedido_oficina)].copy()
        
        cts_agrupados = df_ct_limpo.groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()
        resumo_contratos = pd.merge(df_pendentes_total, cts_agrupados, left_on=NF_CNPJ, right_on='CNPJ', how='left')
        
        resumo_contratos['Status'] = resumo_contratos['Contrato'].apply(lambda x: "⚠️ Vínculo Contratual" if pd.notna(x) else "🚨 PENDÊNCIA CRÍTICA")
        resumo_contratos = resumo_contratos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, 'Contrato', NF_VALOR, 'Status']]

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resumo_painel.to_excel(writer, sheet_name='1. PAINEL (GERAL)', index=False)
            resumo_pedidos.to_excel(writer, sheet_name='2. PEDIDOS (SÓ PENDENTES)', index=False)
            resumo_contratos.to_excel(writer, sheet_name='3. CONTRATOS (FILTRO FINAL)', index=False)
        
        st.success("Auditoria em Funil concluída!")
        st.download_button(label="📥 Baixar Relatório Master Funil", data=output.getvalue(), file_name="AUDITORIA_FUNIL_FINAL.xlsx", mime="application/vnd.ms-excel")
