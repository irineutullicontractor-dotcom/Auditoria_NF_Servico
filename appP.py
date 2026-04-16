import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF - Produto", layout="wide")

st.title("📊 Auditoria Interna NF - Produto")

# --- FUNÇÕES DE PADRONIZAÇÃO TOTAL ---

def limpar_geral(v):
    """Remove nulos, converte para string, tira .0 de números e espaços."""
    if pd.isna(v): return ""
    s = str(v).strip().split('.')[0]
    return "" if s.lower() == "nan" else s

def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def extrair_nf(v):
    """Pega '444/1' e retorna '444'. Também limpa .0 e espaços."""
    if pd.isna(v): return ""
    # Primeiro tira o que vem depois da barra, depois tira o .0 se houver
    s = str(v).split('/')[0].split('.')[0].strip()
    return "" if s.lower() == "nan" else s

# --- PROCESSAMENTO DE ARQUIVOS ---

def estruturar_notas_produtos(file):
    df_bruto = pd.read_excel(file, header=None)
    registros = []
    cnpj_dest, colunas_id, processando = None, None, False

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

def tratar_credores(file):
    df_bruto = pd.read_excel(file, header=None)
    for i in range(min(15, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df = df_bruto.iloc[i+1:].copy()
            df.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            # Extração de código do fornecedor
            def split_forn(val):
                s = str(val).strip()
                return s.split(" - ")[0] if " - " in s else ""
            df['Cód. Fornecedor'] = df['Credor'].apply(split_forn).apply(limpar_geral)
            df['CNPJCPF'] = df['CNPJ/CPF'].apply(limpar_cnpj)
            df['Credor_UP'] = df['Credor'].astype(str).str.upper().strip()
            return df[['Cód. Fornecedor', 'CNPJCPF', 'Credor_UP']]
    return pd.DataFrame()

# --- INTERFACE ---
col1, col2 = st.columns(2)
with col1:
    f_nf = st.file_uploader("1. Relatório de NF's", type=['xlsx'])
    f_forn = st.file_uploader("2. Relatório de Credores", type=['xlsx'])
    f_painel = st.file_uploader("3. Relatório Painel", type=['xlsx'])
with col2:
    f_ped = st.file_uploader("4. Relatório Pedidos", type=['xlsx'])
    f_ct = st.file_uploader("5. Relatório Contrato", type=['xlsx'])

if st.button("🚀 Iniciar Auditoria"):
    if all([f_nf, f_forn, f_painel, f_ped, f_ct]):
        # 1. Cargas e Padronização de Tipos (Strings em Tudo)
        df_nf = estruturar_notas_produtos(f_nf)
        df_forn = tratar_credores(f_forn)
        df_painel = pd.read_excel(f_painel)
        df_ped = pd.read_excel(f_ped)
        
        # Limpeza Crítica das Notas (Aplica na coluna visual e na chave)
        df_nf['Núm/Série'] = df_nf['Núm/Série'].apply(extrair_nf)
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj)
        df_nf['chave'] = df_nf['CNPJ emitente'] + "_" + df_nf['Núm/Série']
        
        # 2. Vínculo com Painel
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.upper().strip()
        painel_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        
        painel_cnpj['nf_limpa'] = painel_cnpj['N° da Nota fiscal'].apply(extrair_nf)
        painel_cnpj['cnpj_limpo'] = painel_cnpj['CNPJCPF'].apply(limpar_cnpj)
        painel_cnpj['chave_p'] = painel_cnpj['cnpj_limpo'] + "_" + painel_cnpj['nf_limpa']
        
        # Cruzamento principal (Aba 1)
        resumo = pd.merge(df_nf, painel_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), 
                          left_on='chave', right_on='chave_p', how='left')
        
        chaves_no_painel = set(painel_cnpj[painel_cnpj['nf_limpa'] != ""]['chave_p'])
        cnpjs_no_painel = set(painel_cnpj['cnpj_limpo'])

        def st_painel(r):
            if r['chave'] in chaves_no_painel: return "✅ NF Lançada"
            if r['CNPJ emitente'] in cnpjs_no_painel: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        resumo['Status'] = resumo.apply(st_painel, axis=1)

        # 3. Vínculo com Pedidos (Aba 2) - Resolvendo o ValueError
        df_ped['cod_forn_ped'] = df_ped['Cód. fornecedor'].apply(limpar_geral)
        # Forçamos o merge de códigos como string
        ped_com_cnpj = pd.merge(df_ped, df_forn[['Cód. Fornecedor', 'CNPJCPF']], 
                                left_on='cod_forn_ped', right_on='Cód. Fornecedor', how='left')
        
        peds_agrup = ped_com_cnpj.groupby('CNPJCPF')['Nº do pedido'].apply(
            lambda x: ", ".join(sorted(set(limpar_geral(v) for v in x if pd.notna(v))))
        ).reset_index()
        
        resumo = pd.merge(resumo, peds_agrup, left_on='CNPJ emitente', right_on='CNPJCPF', how='left')

        # 4. Exportação
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            cols = ['Núm/Série', 'CNPJ emitente', 'Emitente', 'Emissão', 'Valor', 'Status', 'N° da Nota fiscal', 'Nº do pedido']
            resumo[cols].to_excel(writer, sheet_name='Auditoria', index=False)
        
        st.success("Auditoria concluída com sucesso!")
        st.download_button("📥 Baixar Relatório", output.getvalue(), "AUDITORIA_FINAL.xlsx")
