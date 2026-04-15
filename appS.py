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
    if "Cód. Fornecedor" in df_bruto.columns and "Credor" in df_bruto.columns:
        return df_bruto
    
    for i in range(min(10, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            df_header = df_header.loc[:, df_header.columns.notna() & (df_header.columns != 'nan')]
            
            def split_safe(val):
                s = str(val).strip()
                if s == "" or s == "nan": return "", ""
                if " - " in s:
                    parts = s.split(" - ")
                    return parts[0].strip(), " - ".join(parts[1:]).strip()
                return "", s

            res_split = df_header['Credor'].apply(split_safe)
            df_header['Cód. Fornecedor'] = res_split.apply(lambda x: x[0])
            df_header['Fornecedor'] = res_split.apply(lambda x: x[1])
            df_header = df_header.rename(columns={'CNPJ/CPF': 'CNPJCPF'})
            return df_header.dropna(subset=['Credor'])
    return df_bruto

# --- FUNÇÕES DE LIMPEZA ---
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

if st.button("🚀 Processar Auditoria"):
    if all([file_nf, file_forn, file_painel, file_relacao, file_contrato]):
        # Carregamento
        df_nf = pd.read_excel(file_nf) if file_nf.name.endswith('xlsx') else pd.read_csv(file_nf)
        df_forn_raw = pd.read_excel(file_forn, header=None) if file_forn.name.endswith('xlsx') else pd.read_csv(file_forn, header=None)
        df_painel = pd.read_excel(file_painel) if file_painel.name.endswith('xlsx') else pd.read_csv(file_painel)
        df_relacao = pd.read_excel(file_relacao) if file_relacao.name.endswith('xlsx') else pd.read_csv(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None) if file_contrato.name.endswith('xlsx') else pd.read_csv(file_contrato, header=None)

        # Mapeamento de Colunas
        df_forn = transformar_credor_limpo(df_forn_raw)
        
        NF_CNPJ = 'CNPJ Prestador (CNPJ)' if 'CNPJ Prestador (CNPJ)' in df_nf.columns else df_nf.columns[16]
        NF_NUMERO = 'Número NFS-e (nNFSe)'
        NF_FORN = 'Nome Prestador (xNome)'
        NF_DATA = 'Data/Hora Emissão DPS (dhEmi)'
        NF_VALOR = 'Valor do Serviço (vServ) (vServ)'
        
        # Limpezas Iniciais
        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].apply(limpar_cnpj)
        df_nf['nf_limpa'] = df_nf[NF_NUMERO].astype(str).str.strip()
        df_nf['chave_unica'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']
        
        df_forn['CNPJCPF'] = df_forn['CNPJCPF'].apply(limpar_cnpj)
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.strip().str.upper()

        # --- ABA 1: PAINEL ---
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        df_painel['nf_ref_limpa'] = df_painel['N° da Nota fiscal'].apply(extrair_nf)
        
        painel_com_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJCPF'] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        chaves_lancadas_real = set(painel_com_cnpj[painel_com_cnpj['nf_ref_limpa'] != ""]['chave_p'].unique())
        cnpjs_no_painel = set(painel_com_cnpj['CNPJCPF'].unique())

        resumo_painel = pd.merge(df_nf, painel_com_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), left_on='chave_unica', right_on='chave_p', how='left')
        
        def definir_status_painel(r):
            if r['chave_unica'] in chaves_lancadas_real: return "✅ NF Lançada"
            if r[NF_CNPJ] in cnpjs_no_painel: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_painel['Status'] = resumo_painel.apply(definir_status_painel, axis=1)
        aba1_final = resumo_painel[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, 'N° da Nota fiscal', 'Status']]

        # --- ABA 2: PEDIDOS ---
        df_relacao['Cód. fornecedor'] = df_relacao['Cód. fornecedor'].apply(limpar_cod)
        rel_com_cnpj = pd.merge(df_relacao, df_forn[['Cód. Fornecedor', 'CNPJCPF']], left_on='Cód. fornecedor', right_on='Cód. Fornecedor', how='left')
        peds_agrupados = rel_com_cnpj.groupby('CNPJCPF')['Nº do pedido'].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()

        resumo_pedidos = pd.merge(resumo_painel, peds_agrupados, left_on=NF_CNPJ, right_on='CNPJCPF', how='left')
        cnpjs_com_pedido = set(peds_agrupados['CNPJCPF'].unique())

        def status_pedidos(r):
            if r['chave_unica'] in chaves_lancadas_real: return "✅ Resolvido Painel"
            if r[NF_CNPJ] in cnpjs_com_pedido or r['Status'] == "⚠️ Para Verificação": return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_pedidos['Status_Ped'] = resumo_pedidos.apply(status_pedidos, axis=1)
        aba2_final = resumo_pedidos[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, 'N° da Nota fiscal', 'Nº do pedido', 'Status_Ped']].rename(columns={'Status_Ped': 'Status', 'Nº do pedido': 'Pedido'})

        # --- ABA 3: CONTRATO (LÓGICA ESPELHADA NO MODELO FUNCIONAL) ---
        registros_ct = []
        item_atual = {'Contrato': None, 'CNPJ': None}
        
        for i in range(len(df_bruto_ct)):
            linha = df_bruto_ct.iloc[i]
            col_a = str(linha[0]).strip() if pd.notna(linha[0]) else ""
            col_d = linha[3] if pd.notna(linha[3]) else "" # Pega o valor bruto (ex: CT/500)
            
            # Se encontrar "Contrato" na coluna A, salva o número na variável temporária
            if col_a == "Contrato":
                item_atual['Contrato'] = str(col_d).strip()
            
            # Se encontrar "CNPJ" na coluna A, vincula ao contrato atual e salva na lista
            elif col_a == "CNPJ" and item_atual['Contrato']:
                item_atual['CNPJ'] = limpar_cnpj(col_d)
                registros_ct.append(item_atual.copy())
                # Não limpamos o contrato aqui porque pode haver mais de um CNPJ para o mesmo contrato em alguns relatórios

        df_ct = pd.DataFrame(registros_ct).dropna(subset=['CNPJ'])
        # Agrupa caso o mesmo CNPJ tenha vários contratos
        cts_agrupados = df_ct.groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(set(x.astype(str).unique()))).reset_index()

        # Faz o cruzamento com o resumo anterior
        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on=NF_CNPJ, right_on='CNPJ', how='left')
        
        def status_contratos(r):
            if r['chave_unica'] in chaves_lancadas_real: return "✅ Resolvido Painel"
            if pd.notna(r['Contrato']) and str(r['Contrato']).strip() != "": return "📄 Vínculo Contratual"
            return r['Status_Ped']

        resumo_contratos['Status_CT'] = resumo_contratos.apply(status_contratos, axis=1)
        
        aba3_final = resumo_contratos[[
            NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, 
            'N° da Nota fiscal', 'Nº do pedido', 'Contrato', 'Status_CT'
        ]].rename(columns={'Status_CT': 'Status', 'Nº do pedido': 'Pedido'})

        # Exportação
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1_final.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2_final.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3_final.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Tudo pronto! Relatório de Auditoria gerado com sucesso.")
        st.download_button(label="📥 Baixar Auditoria", data=output.getvalue(), file_name="AUDITORIA_NF.xlsx")
