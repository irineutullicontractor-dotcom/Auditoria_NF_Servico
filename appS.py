import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Master: Fluxo Inteligente", layout="wide")

st.title("📊 Auditoria Master: Painel -> Pedidos -> Contratos")
st.markdown("""
### Regras do Fluxo:
1. **Painel:** Identifica o que está lançado e o que tem histórico de pedido no painel.
2. **Pedidos:** Herda o status do Painel e adiciona a verificação da Oficina/Pedidos.
3. **Contrato:** Consolida tudo, mantendo os alertas anteriores e adicionando o Vínculo Contratual.
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
    if file.name.endswith('.csv'): return pd.read_csv(file)
    return pd.read_excel(file, header=header)

if st.button("🚀 Processar Fluxo Consolidado"):
    if not all([file_nf, file_forn, file_painel, file_relacao, file_contrato]):
        st.error("Por favor, carregue os 5 arquivos.")
    else:
        # Carregamento e Limpeza de Colunas
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
        NF_CNPJ = encontrar_coluna(df_nf, ['CNPJ Prestador (CNPJ)', 'Prestador (CNPJ)', 'Prestador (CNPJ / CPF)', 'CNPJ'])
        NF_NUMERO = encontrar_coluna(df_nf, ['Número NFS-e (nNFSe)', 'Número (nNFSe)', 'nNFSe'])
        NF_FORN = encontrar_coluna(df_nf, ['Nome Prestador (xNome)', 'Prestador (xNome)'])
        NF_DATA = encontrar_coluna(df_nf, ['Data/Hora Emissão DPS (dhEmi)', 'Data da Emissão (dhEmi)'])
        NF_VALOR = encontrar_coluna(df_nf, ['Valor do Serviço (vServ) (vServ)', 'Valor Serviço (vServ)'])
        
        PED_FORN_REL = encontrar_coluna(df_relacao, ['Cód. fornecedor', 'Cód. Fornecedor', 'Fornecedor'])
        PED_NUM_REL = encontrar_coluna(df_relacao, ['Nº do pedido', 'N° do Pedido', 'Pedido'])
        
        PED_FORN_PAINEL, PED_NUM_PAINEL, PED_NF_REF = 'Fornecedor', 'N° do Pedido', 'N° da Nota fiscal'
        FORN_COD, FORN_CNPJ, FORN_CRED = 'Cód. Fornecedor', 'CNPJCPF', 'Credor'

        # Funções de limpeza
        def limpar_cnpj(v):
            num = "".join(filter(str.isdigit, str(v)))
            return num.zfill(14) if len(num) > 11 else num.zfill(11)

        def extrair_nf(v):
            if pd.isna(v) or v == "": return ""
            return "".join(filter(str.isdigit, str(v).split('/')[-1])).strip()

        # Tratamento de dados
        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf[NF_NUMERO].astype(str).str.strip()
        df_forn[FORN_CNPJ] = df_forn[FORN_CNPJ].apply(limpar_cnpj)
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

       # --- PADRONIZAÇÕES ---
        def limpar_cnpj(v):
            if pd.isna(v): return ""
            num = "".join(filter(str.isdigit, str(v)))
            return num.zfill(14) if len(num) > 11 else num.zfill(11)

        def limpar_cod(v):
            if pd.isna(v): return ""
            # Transforma em string, remove o .0 se for float, limpa espaços e zeros à esquerda
            return str(v).split('.')[0].strip().lstrip('0')

        def extrair_nf(v):
            if pd.isna(v) or v == "": return ""
            return "".join(filter(str.isdigit, str(v).split('/')[-1])).strip()

        # Aplicação das limpezas com tratamento de erro
        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf[NF_NUMERO].astype(str).str.strip()
        df_forn[FORN_CNPJ] = df_forn[FORN_CNPJ].apply(limpar_cnpj)
        df_forn[FORN_COD] = df_forn[FORN_COD].apply(limpar_cod)
        df_forn[FORN_CRED] = df_forn[FORN_CRED].str.strip().str.upper()

        # =================================================================
        # ABA 1: PAINEL
        # =================================================================
        # (Mantém o código anterior da Aba 1 aqui...)
        df_painel[PED_FORN_PAINEL] = df_painel[PED_FORN_PAINEL].str.strip().str.upper()
        df_painel['nf_extraida'] = df_painel[PED_NF_REF].apply(extrair_nf)
        painel_com_cnpj = pd.merge(df_painel, df_forn[[FORN_CRED, FORN_CNPJ]], left_on=PED_FORN_PAINEL, right_on=FORN_CRED, how='left')
        
        df_nf['chave'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']
        painel_com_cnpj['chave'] = painel_com_cnpj[FORN_CNPJ] + "_" + painel_com_cnpj['nf_extraida']
        
        nfs_lancadas = df_nf[df_nf['chave'].isin(painel_com_cnpj['chave'])]['nf_limpa'].unique()
        cnpjs_com_painel = painel_com_cnpj[FORN_CNPJ].unique()

        resumo_painel = df_nf.copy()
        resumo_painel['Status'] = resumo_painel.apply(lambda r: "✅ NF Lançada" if r['nf_limpa'] in nfs_lancadas else ("⚠️ Para Verificação" if r[NF_CNPJ] in cnpjs_com_painel else "❌ Sem Histórico"), axis=1)

        # =================================================================
        # ABA 2: PEDIDOS (CORREÇÃO DO ERRO AQUI)
        # =================================================================
        # Limpamos o código do fornecedor no arquivo de Relação/Oficina
        if PED_FORN_REL in df_relacao.columns:
            df_relacao[PED_FORN_REL] = df_relacao[PED_FORN_REL].apply(limpar_cod)
            
            # Cruzamento com fornecedores para pegar o CNPJ
            forn_com_cod = pd.merge(df_relacao, df_forn[[FORN_COD, FORN_CNPJ]], left_on=PED_FORN_REL, right_on=FORN_COD, how='left')
            cnpjs_com_oficina = forn_com_cod[FORN_CNPJ].dropna().unique()
        else:
            cnpjs_com_oficina = []

        resumo_pedidos = resumo_painel.copy()
        def status_pedidos(r):
            if r['Status'] == "✅ NF Lançada": return "✅ Resolvido Painel"
            if r['Status'] == "⚠️ Para Verificação" or r[NF_CNPJ] in cnpjs_com_oficina:
                return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_pedidos['Status'] = resumo_pedidos.apply(status_pedidos, axis=1)

        # =================================================================
        # ABA 3: CONTRATOS (CONSOLIDAÇÃO FINAL)
        # =================================================================
        cnpjs_com_contrato = df_ct_limpo['CNPJ'].unique()

        resumo_contratos = resumo_pedidos.copy()
        def status_contratos(r):
            if r['Status'] == "✅ Resolvido Painel": return "✅ Resolvido Painel"
            # Se tem contrato, ganha o status prioritário
            if r[NF_CNPJ] in cnpjs_com_contrato: return "📄 Vínculo Contratual"
            # Mantém o que veio das abas anteriores
            return r['Status']

        resumo_contratos['Status'] = resumo_contratos.apply(status_contratos, axis=1)

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resumo_painel.to_excel(writer, sheet_name='1. PAINEL', index=False)
            resumo_pedidos.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            resumo_contratos.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Relatório consolidado com sucesso!")
        st.download_button(label="📥 Baixar Auditoria Final", data=output.getvalue(), file_name="AUDITORIA_CONSOLIDADA.xlsx")
