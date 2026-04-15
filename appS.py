import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Master: Funil de Colunas", layout="wide")

st.title("📊 Auditoria Master: Fluxo de Dados Consolidado")

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
    if file.name.endswith('.csv'): return pd.read_csv(file)
    return pd.read_excel(file, header=header)

if st.button("🚀 Processar Auditoria"):
    if not all([file_nf, file_forn, file_painel, file_relacao, file_contrato]):
        st.error("Por favor, carregue os 5 arquivos.")
    else:
        # Carregamento
        df_nf = carregar(file_nf)
        df_forn = carregar(file_forn)
        df_painel = carregar(file_painel)
        df_relacao = carregar(file_relacao)
        df_bruto_ct = carregar(file_contrato, header=None)

        def encontrar_coluna(df, opcoes):
            df.columns = df.columns.str.strip()
            for opt in opcoes:
                if opt in df.columns: return opt
            return None

        # Mapeamentos
        NF_CNPJ = encontrar_coluna(df_nf, ['CNPJ Prestador (CNPJ)', 'Prestador (CNPJ)', 'Prestador (CNPJ / CPF)'])
        NF_NUMERO = encontrar_coluna(df_nf, ['Número NFS-e (nNFSe)', 'Número (nNFSe)'])
        NF_FORN = encontrar_coluna(df_nf, ['Nome Prestador (xNome)', 'Prestador (xNome)'])
        NF_DATA = encontrar_coluna(df_nf, ['Data/Hora Emissão DPS (dhEmi)', 'Data da Emissão (dhEmi)'])
        NF_VALOR = encontrar_coluna(df_nf, ['Valor do Serviço (vServ) (vServ)', 'Valor Serviço (vServ)'])
        
        PED_FORN_REL = encontrar_coluna(df_relacao, ['Cód. fornecedor', 'Cód. Fornecedor'])
        PED_NUM_REL = encontrar_coluna(df_relacao, ['Nº do pedido', 'N° do Pedido'])
        
        PED_FORN_PAINEL, PED_NUM_PAINEL, PED_NF_REF = 'Fornecedor', 'N° do Pedido', 'N° da Nota fiscal'
        FORN_COD, FORN_CNPJ, FORN_CRED = 'Cód. Fornecedor', 'CNPJCPF', 'Credor'

        # Limpezas
        def limpar_cnpj(v):
            if pd.isna(v): return ""
            num = "".join(filter(str.isdigit, str(v)))
            return num.zfill(14) if len(num) > 11 else num.zfill(11)

        def limpar_cod(v):
            if pd.isna(v): return ""
            return str(v).split('.')[0].strip().lstrip('0')

        def extrair_nf(v):
            if pd.isna(v) or v == "": return ""
            return "".join(filter(str.isdigit, str(v).split('/')[-1])).strip()

        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf[NF_NUMERO].astype(str).str.strip()
        df_forn[FORN_CNPJ] = df_forn[FORN_CNPJ].apply(limpar_cnpj)
        df_forn[FORN_COD] = df_forn[FORN_COD].apply(limpar_cod)
        df_forn[FORN_CRED] = df_forn[FORN_CRED].str.strip().str.upper()

        # Limpeza Contrato Bruto
        registros_ct = []
        item_atual = {}
        for i, row in df_bruto_ct.iterrows():
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

        # --- ABA 1: PAINEL ---
        df_painel[PED_FORN_PAINEL] = df_painel[PED_FORN_PAINEL].str.strip().str.upper()
        df_painel['nf_ref_limpa'] = df_painel[PED_NF_REF].apply(extrair_nf)
        painel_com_cnpj = pd.merge(df_painel, df_forn[[FORN_CRED, FORN_CNPJ]], left_on=PED_FORN_PAINEL, right_on=FORN_CRED, how='left')
        
        df_nf['chave'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']
        painel_com_cnpj['chave_p'] = painel_com_cnpj[FORN_CNPJ] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        # Guardamos a info do painel (N° da Nota fiscal)
        painel_info = painel_com_cnpj[['chave_p', PED_NF_REF]].drop_duplicates('chave_p')
        resumo_painel = pd.merge(df_nf, painel_info, left_on='chave', right_on='chave_p', how='left')
        
        nfs_lancadas = resumo_painel[resumo_painel['chave_p'].notna()]['nf_limpa'].unique()
        cnpjs_no_painel = painel_com_cnpj[FORN_CNPJ].unique()

        def status_painel(r):
            if pd.notna(r['chave_p']): return "✅ NF Lançada"
            if r[NF_CNPJ] in cnpjs_no_painel: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_painel['Status'] = resumo_painel.apply(status_painel, axis=1)
        # ABA 1: Inclui "N° da Nota fiscal"
        aba1_final = resumo_painel[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, PED_NF_REF, 'Status']]
        aba1_final = aba1_final.rename(columns={PED_NF_REF: 'N° da Nota fiscal'})

        # --- ABA 2: PEDIDOS ---
        df_relacao[PED_FORN_REL] = df_relacao[PED_FORN_REL].apply(limpar_cod)
        rel_com_cnpj = pd.merge(df_relacao, df_forn[[FORN_COD, FORN_CNPJ]], left_on=PED_FORN_REL, right_on=FORN_COD, how='left')
        peds_oficina = rel_com_cnpj.groupby(FORN_CNPJ)[PED_NUM_REL].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()

        resumo_pedidos = pd.merge(resumo_painel, peds_oficina, left_on=NF_CNPJ, right_on=FORN_CNPJ, how='left')
        cnpjs_na_oficina = peds_oficina[FORN_CNPJ].unique()

        def status_pedidos(r):
            if r['nf_limpa'] in nfs_lancadas: return "✅ Resolvido Painel"
            if r['Status'] == "⚠️ Para Verificação" or r[NF_CNPJ] in cnpjs_na_oficina: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_pedidos['Status_Ped'] = resumo_pedidos.apply(status_pedidos, axis=1)
        # ABA 2: Carrega "N° da Nota fiscal" e adiciona "Pedido"
        aba2_final = resumo_pedidos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, PED_NF_REF, PED_NUM_REL, 'Status_Ped']]
        aba2_final = aba2_final.rename(columns={'Status_Ped': 'Status', PED_NF_REF: 'N° da Nota fiscal', PED_NUM_REL: 'Pedido'})

        # --- ABA 3: CONTRATO ---
        cts_agrupados = df_ct_limpo.groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()
        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on=NF_CNPJ, right_on='CNPJ', how='left')
        cnpjs_com_ct = cts_agrupados['CNPJ'].unique()

        def status_contratos(r):
            if r['nf_limpa'] in nfs_lancadas: return "✅ Resolvido Painel"
            if r[NF_CNPJ] in cnpjs_com_ct: return "📄 Vínculo Contratual"
            return r['Status_Ped']

        resumo_contratos['Status_CT'] = resumo_contratos.apply(status_contratos, axis=1)
        # ABA 3: Carrega todas (N° NF + Pedido) e adiciona "Contrato"
        aba3_final = resumo_contratos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, PED_NF_REF, PED_NUM_REL, 'Contrato', 'Status_CT']]
        aba3_final = aba3_final.rename(columns={'Status_CT': 'Status', PED_NF_REF: 'N° da Nota fiscal', PED_NUM_REL: 'Pedido'})

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1_final.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2_final.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3_final.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Relatório concluído com herança completa de colunas!")
        st.download_button(label="📥 Baixar Auditoria Consolidada", data=output.getvalue(), file_name="AUDITORIA_MASTER.xlsx")
