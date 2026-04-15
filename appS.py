import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF", layout="wide")

st.title("📊 Auditoria Interna NF")

# --- UPLOAD DOS 5 FICHEIROS ---
col1, col2 = st.columns(2)
with col1:
    file_nf = st.file_uploader("1. Relatório de NF's", type=['xlsx', 'csv'])
    file_forn = st.file_uploader("2. Relatório de Credores", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato", type=['xlsx', 'csv'])

def carregar(file, header=0):
    if file is None: return None
    if file.name.endswith('.csv'): return pd.read_csv(file)
    return pd.read_excel(file, header=header)

# --- FUNÇÃO DE TRANSFORMAÇÃO DE CREDORES (BRUTO PARA LIMPO) ---
def transformar_credor_limpo(df_bruto):
    # Se já tiver a coluna 'Cód. Fornecedor', o arquivo já está no padrão limpo
    if "Cód. Fornecedor" in df_bruto.columns:
        return df_bruto
    
    # Localiza o cabeçalho real no arquivo bruto (procura a linha que contém 'Credor')
    for i in range(len(df_bruto)):
        # Verifica se a palavra 'Credor' está em alguma célula da linha
        if 'Credor' in df_bruto.iloc[i].astype(str).values:
            # Pega os dados a partir da linha seguinte ao cabeçalho encontrado
            df_limpo = df_bruto.iloc[i+1:].copy()
            df_limpo.columns = df_bruto.iloc[i].values
            
            # Remove colunas vazias (Unamed)
            df_limpo = df_limpo.loc[:, df_limpo.columns.notna()]
            
            # Padroniza nomes de colunas
            df_limpo = df_limpo.rename(columns={'Credor': 'Credor', 'CNPJ/CPF': 'CNPJCPF'})
            
            # --- RECONSTRUÇÃO DAS COLUNAS CONFORME MODELO MISTO ---
            # 1. 'Credor' original é mantido.
            # 2. 'Cód. Fornecedor': Extrai o número antes do " - "
            df_limpo['Cód. Fornecedor'] = df_limpo['Credor'].astype(str).str.split(' - ').str[0].str.strip()
            
            # 3. 'Fornecedor': Extrai o nome após o " - "
            df_limpo['Fornecedor'] = df_limpo['Credor'].astype(str).str.split(' - ').str[1:].apply(lambda x: " - ".join(x)).str.strip()
            
            # 4. 'CNPJCPF': Já renomeado acima.
            
            # Reordena para ficar idêntico ao arquivo fornecedores-misto.xlsx
            colunas_finais = ['Cód. Fornecedor', 'Fornecedor', 'Credor', 'CNPJCPF']
            # Filtra apenas o que existe para não dar erro
            colunas_existentes = [c for c in colunas_finais if c in df_limpo.columns]
            
            return df_limpo[colunas_existentes].dropna(subset=['Credor'])
            
    return df_bruto

if st.button("🚀 Processar Auditoria"):
    if not all([file_nf, file_forn, file_painel, file_relacao, file_contrato]):
        st.error("Por favor, carregue os 5 arquivos.")
    else:
        # Carregamento (usando header=None para os que precisam de tratamento manual)
        df_nf = carregar(file_nf)
        df_forn_raw = carregar(file_forn, header=None)
        df_painel = carregar(file_painel)
        df_relacao = carregar(file_relacao)
        df_bruto_ct = carregar(file_contrato, header=None)

        # Transforma o arquivo de Credores bruto no formato Misto/Limpo automaticamente
        df_forn = transformar_credor_limpo(df_forn_raw)

        def encontrar_coluna(df, opcoes):
            df.columns = [str(c).strip() for c in df.columns]
            for opt in opcoes:
                if opt in df.columns: return opt
            return None

        # Mapeamentos (usando os nomes padronizados pela limpeza acima)
        FORN_COD, FORN_CNPJ, FORN_CRED = 'Cód. Fornecedor', 'CNPJCPF', 'Credor'
        
        NF_CNPJ = encontrar_coluna(df_nf, ['CNPJ Prestador (CNPJ)', 'Prestador (CNPJ)', 'Prestador (CNPJ / CPF)'])
        NF_NUMERO = encontrar_coluna(df_nf, ['Número NFS-e (nNFSe)', 'Número (nNFSe)'])
        NF_FORN = encontrar_coluna(df_nf, ['Nome Prestador (xNome)', 'Prestador (xNome)'])
        NF_DATA = encontrar_coluna(df_nf, ['Data/Hora Emissão DPS (dhEmi)', 'Data da Emissão (dhEmi)'])
        NF_VALOR = encontrar_coluna(df_nf, ['Valor do Serviço (vServ) (vServ)', 'Valor Serviço (vServ)'])
        
        PED_FORN_REL = encontrar_coluna(df_relacao, ['Cód. fornecedor', 'Cód. Fornecedor'])
        PED_NUM_REL = encontrar_coluna(df_relacao, ['Nº do pedido', 'N° do Pedido'])
        
        PED_FORN_PAINEL, PED_NUM_PAINEL, PED_NF_REF = 'Fornecedor', 'N° do Pedido', 'N° da Nota fiscal'

        # Funções de Limpeza de String
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

        # Aplicando Tratamentos
        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf[NF_NUMERO].astype(str).str.strip()
        
        df_forn[FORN_CNPJ] = df_forn[FORN_CNPJ].apply(limpar_cnpj)
        df_forn[FORN_COD] = df_forn[FORN_COD].apply(limpar_cod)
        # Importante: O nome do fornecedor no painel deve bater com o nome na coluna 'Credor' (que é Código - Nome)
        df_forn[FORN_CRED] = df_forn[FORN_CRED].astype(str).str.strip().str.upper()

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
        df_painel[PED_FORN_PAINEL] = df_painel[PED_FORN_PAINEL].astype(str).str.strip().str.upper()
        df_painel['nf_ref_limpa'] = df_painel[PED_NF_REF].apply(extrair_nf)
        
        # Merge de fornecedores com Painel usando o nome completo (Credor)
        painel_com_cnpj = pd.merge(df_painel, df_forn[[FORN_CRED, FORN_CNPJ]], left_on=PED_FORN_PAINEL, right_on=FORN_CRED, how='left')
        
        df_nf['chave'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']
        painel_com_cnpj['chave_p'] = painel_com_cnpj[FORN_CNPJ] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        painel_info = painel_com_cnpj[['chave_p', PED_NF_REF]].drop_duplicates('chave_p')
        resumo_painel = pd.merge(df_nf, painel_info, left_on='chave', right_on='chave_p', how='left')
        
        nfs_lancadas = resumo_painel[resumo_painel['chave_p'].notna()]['nf_limpa'].unique()
        cnpjs_no_painel = painel_com_cnpj[FORN_CNPJ].unique()

        def status_painel(r):
            if pd.notna(r['chave_p']): return "✅ NF Lançada"
            if r[NF_CNPJ] in cnpjs_no_painel: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_painel['Status'] = resumo_painel.apply(status_painel, axis=1)
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
        aba3_final = resumo_contratos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, PED_NF_REF, PED_NUM_REL, 'Contrato', 'Status_CT']]
        aba3_final = aba3_final.rename(columns={'Status_CT': 'Status', PED_NF_REF: 'N° da Nota fiscal', PED_NUM_REL: 'Pedido'})

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1_final.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2_final.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3_final.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Relatório gerado! O arquivo de credores foi limpo e transformado no padrão misto internamente.")
        st.download_button(label="📥 Baixar Auditoria Consolidada", data=output.getvalue(), file_name="AUDITORIA_NF_SERVICO.xlsx")
