import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF", layout="wide")

st.title("📊 Auditoria Interna NF")
st.markdown("""

### Instruções de uso:
1. Carregue o relatório de **NF's** - 1 por período.
2. Carregue o relatório de **Credores**.
3. Carregue o relatório do **Painel** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
2. Carregue o relatório de **Pedidos** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
4. Carregue o relatório de **Contratos** - Puxar relatório de 01/01/2020 até a data vigente.
""")

# --- UPLOAD DOS 5 FICHEIROS ---
col1, col2 = st.columns(2)
with col1:
    file_nf = st.file_uploader("1. Relatório de NF's - Fornecido a cada 10 dias no servidor.", type=['xlsx', 'csv'])
    file_forn = st.file_uploader("2. Relatório de Credores - Home / Mais Opções / Apoio / Relatórios / Pessoas / Credores.", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel - Home / Suprimentos / Compras / Painel de Compras (Novo).", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos - Home / Suprimentos / Compras / Relatórios / Pedidos de compra / Relação de Pedidos de Compra (Novo).", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato - Home / Suprimentos / Contratos e Medições / Relatórios / Contratos / Emissão de Contratos.", type=['xlsx', 'csv'])

def carregar(file, header=0):
    if file is None: return None
    if file.name.endswith('.csv'): return pd.read_csv(file)
    return pd.read_excel(file, header=header)

# --- FUNÇÃO ROBUSTA DE TRANSFORMAÇÃO DE CREDORES ---
def transformar_credor_limpo(df_bruto):
    # 1. Verifica se já está no formato limpo (misto) pelas colunas
    if "Cód. Fornecedor" in df_bruto.columns and "Credor" in df_bruto.columns:
        return df_bruto
    
    # 2. Verifica se a primeira linha contém os cabeçalhos (caso o header=None tenha sido usado)
    if not df_bruto.empty:
        first_row = [str(x).strip() for x in df_bruto.iloc[0].values]
        if "Cód. Fornecedor" in first_row:
            df_ajustado = df_bruto.iloc[1:].copy()
            df_ajustado.columns = first_row
            return df_ajustado

    # 3. Formato Bruto: Procura a linha que contém "Credor" e "CNPJ/CPF"
    for i in range(len(df_bruto)):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            
            # Remove colunas fantasmas (nan)
            df_header = df_header.loc[:, df_header.columns.notna() & (df_header.columns != 'nan')]
            
            # Função para separar "1 - Empresa" em ("1", "Empresa")
            def split_safe(val):
                s = str(val).strip()
                if s == "" or s == "nan": return "", ""
                if " - " in s:
                    parts = s.split(" - ")
                    return parts[0].strip(), " - ".join(parts[1:]).strip()
                return "", s

            # Criação das colunas idênticas ao arquivo "misto"
            res_split = df_header['Credor'].apply(split_safe)
            df_header['Cód. Fornecedor'] = res_split.apply(lambda x: x[0])
            df_header['Fornecedor'] = res_split.apply(lambda x: x[1])
            df_header = df_header.rename(columns={'CNPJ/CPF': 'CNPJCPF'})
            
            # Define a ordem exata das colunas do arquivo misto
            cols_misto = ['Cód. Fornecedor', 'Fornecedor', 'Credor', 'CNPJCPF']
            df_final = df_header[[c for c in cols_misto if c in df_header.columns]].copy()
            
            # Remove linhas onde o Credor é inválido
            df_final = df_final[df_final['Credor'].astype(str).str.strip().str.lower() != 'nan']
            return df_final.dropna(subset=['Credor'])
            
    return df_bruto

if st.button("🚀 Processar Auditoria"):
    if not all([file_nf, file_forn, file_painel, file_relacao, file_contrato]):
        st.error("Por favor, carregue os 5 arquivos.")
    else:
        # Carregamento (Credores e Contratos carregados sem header para limpeza manual)
        df_nf = carregar(file_nf)
        df_forn_raw = carregar(file_forn, header=None)
        df_painel = carregar(file_painel)
        df_relacao = carregar(file_relacao)
        df_bruto_ct = carregar(file_contrato, header=None)

        # Transformação automática do Credor
        df_forn = transformar_credor_limpo(df_forn_raw)

        def encontrar_coluna(df, opcoes):
            df.columns = [str(c).strip() for c in df.columns]
            for opt in opcoes:
                if opt in df.columns: return opt
            return None

        # Definição das colunas de referência
        FORN_COD, FORN_CNPJ, FORN_CRED = 'Cód. Fornecedor', 'CNPJCPF', 'Credor'
        
        NF_CNPJ = encontrar_coluna(df_nf, ['CNPJ Prestador (CNPJ)', 'Prestador (CNPJ)', 'Prestador (CNPJ / CPF)'])
        NF_NUMERO = encontrar_coluna(df_nf, ['Número NFS-e (nNFSe)', 'Número (nNFSe)'])
        NF_FORN = encontrar_coluna(df_nf, ['Nome Prestador (xNome)', 'Prestador (xNome)'])
        NF_DATA = encontrar_coluna(df_nf, ['Data/Hora Emissão DPS (dhEmi)', 'Data da Emissão (dhEmi)'])
        NF_VALOR = encontrar_coluna(df_nf, ['Valor do Serviço (vServ) (vServ)', 'Valor Serviço (vServ)'])
        
        PED_FORN_REL = encontrar_coluna(df_relacao, ['Cód. fornecedor', 'Cód. Fornecedor'])
        PED_NUM_REL = encontrar_coluna(df_relacao, ['Nº do pedido', 'N° do Pedido'])
        PED_FORN_PAINEL, PED_NF_REF = 'Fornecedor', 'N° da Nota fiscal'

        # Funções de limpeza
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

        # Aplicação das limpezas
        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf[NF_NUMERO].astype(str).str.strip()
        df_forn[FORN_CNPJ] = df_forn[FORN_CNPJ].apply(limpar_cnpj)
        df_forn[FORN_COD] = df_forn[FORN_COD].apply(limpar_cod)
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
        
        painel_com_cnpj = pd.merge(df_painel, df_forn[[FORN_CRED, FORN_CNPJ]], left_on=PED_FORN_PAINEL, right_on=FORN_CRED, how='left')
        df_nf['chave'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']
        painel_com_cnpj['chave_p'] = painel_com_cnpj[FORN_CNPJ] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        painel_info = painel_com_cnpj[['chave_p', PED_NF_REF]].drop_duplicates('chave_p')
        resumo_painel = pd.merge(df_nf, painel_info, left_on='chave', right_on='chave_p', how='left')
        
        nfs_lancadas = resumo_painel[resumo_painel['chave_p'].notna()]['nf_limpa'].unique()
        cnpjs_no_painel = painel_com_cnpj[FORN_CNPJ].unique()

        resumo_painel['Status'] = resumo_painel.apply(lambda r: "✅ NF Lançada" if pd.notna(r['chave_p']) else ("⚠️ Para Verificação" if r[NF_CNPJ] in cnpjs_no_painel else "❌ Sem Histórico"), axis=1)
        aba1_final = resumo_painel[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, PED_NF_REF, 'Status']].rename(columns={PED_NF_REF: 'N° da Nota fiscal'})

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
        aba2_final = resumo_pedidos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, PED_NF_REF, PED_NUM_REL, 'Status_Ped']].rename(columns={'Status_Ped': 'Status', PED_NF_REF: 'N° da Nota fiscal', PED_NUM_REL: 'Pedido'})

        # --- ABA 3: CONTRATO ---
        cts_agrupados = df_ct_limpo.groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()
        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on=NF_CNPJ, right_on='CNPJ', how='left')
        cnpjs_com_ct = cts_agrupados['CNPJ'].unique()

        def status_contratos(r):
            if r['nf_limpa'] in nfs_lancadas: return "✅ Resolvido Painel"
            if r[NF_CNPJ] in cnpjs_com_ct: return "📄 Vínculo Contratual"
            return r['Status_Ped']

        resumo_contratos['Status_CT'] = resumo_contratos.apply(status_contratos, axis=1)
        aba3_final = resumo_contratos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, PED_NF_REF, PED_NUM_REL, 'Contrato', 'Status_CT']].rename(columns={'Status_CT': 'Status', PED_NF_REF: 'N° da Nota fiscal', PED_NUM_REL: 'Pedido'})

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1_final.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2_final.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3_final.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Relatório gerado com sucesso! A limpeza dos credores foi feita automaticamente.")
        st.download_button(label="📥 Baixar Auditoria Consolidada", data=output.getvalue(), file_name="AUDITORIA_NF_SERVICO.xlsx")
