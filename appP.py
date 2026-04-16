import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF - Produto", layout="wide")

st.title("📊 Auditoria Interna NF - Produto")
st.markdown("""
### Instruções de uso:
1. Carregue o relatório de **NF's** - Puxar relatório do mês vigente.
2. Carregue o relatório de **Credores**.
3. Carregue o relatório do **Painel** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
4. Carregue o relatório de **Pedidos** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
5. Carregue o relatório de **Contratos** - Puxar relatório de 01/01/2020 até a data vigente.
""")

# --- FUNÇÕES DE APOIO ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def limpar_cod(v):
    if pd.isna(v): return ""
    return str(v).split('.')[0].strip().lstrip('0')

def extrair_nf_produto(v):
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    # Pega o que vem antes da barra e remove zeros à esquerda
    num_parte = str(v).split('/')[0].split('-')[0].strip()
    return "".join(filter(str.isdigit, num_parte)).lstrip('0')

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
                if " - " in s:
                    parts = s.split(" - ")
                    return parts[0].strip(), " - ".join(parts[1:]).strip()
                return "", s
            res_split = df_header['Credor'].apply(split_safe)
            df_header['Cód. Fornecedor_Limpo'] = res_split.apply(lambda x: x[0])
            df_header['Fornecedor_Nome'] = res_split.apply(lambda x: x[1])
            return df_header.rename(columns={'CNPJ/CPF': 'CNPJCPF'})
    return df_bruto

# --- UPLOAD ---
col1, col2 = st.columns(2)
with col1:
    file_nf_prod = st.file_uploader("1. Relatório de NF's - Home / Notas Fiscais / Recepção de NF-e / Relatórios / Notas Fiscais Recebidas.", type=['xlsx'])
    file_forn = st.file_uploader("2. Relatório de Credores - Home / Mais Opções / Apoio / Relatórios / Pessoas / Credores.", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel - Home / Suprimentos / Compras / Painel de Compras (Novo).", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos - Home / Suprimentos / Compras / Relatórios / Pedidos de compra / Relação de Pedidos de Compra (Novo).", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato - Home / Suprimentos / Contratos e Medições / Relatórios / Contratos / Emissão de Contratos.", type=['xlsx', 'csv'])

if st.button("🚀 Iniciar Auditoria"):
    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato]):
        
        # 1. Carregar dados
        df_nf = estruturar_notas_produtos_interno(file_nf_prod)
        df_forn = transformar_credor_limpo(pd.read_excel(file_forn, header=None))
        df_painel = pd.read_excel(file_painel)
        df_relacao = pd.read_excel(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None)

        # 2. Padronizar NF e Credores
        df_nf['CNPJ_EMIT_LIMPO'] = df_nf['CNPJ emitente'].apply(limpar_cnpj)
        df_nf['NF_PURA'] = df_nf['Núm/Série'].apply(extrair_nf_produto)
        df_nf['chave_unica'] = df_nf['CNPJ_EMIT_LIMPO'] + "_" + df_nf['NF_PURA']

        df_forn['CNPJ_FORN_LIMPO'] = df_forn['CNPJCPF'].apply(limpar_cnpj)
        df_forn['FORN_UP'] = df_forn['Fornecedor_Nome'].astype(str).str.strip().str.upper()

        # 3. Cruzamento Painel
        df_painel['NF_PAINEL_PURA'] = df_painel['N° da Nota fiscal'].apply(extrair_nf_produto)
        df_painel['FORN_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        
        painel_com_cnpj = pd.merge(df_painel, df_forn[['FORN_UP', 'CNPJ_FORN_LIMPO']], on='FORN_UP', how='left')
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJ_FORN_LIMPO'] + "_" + painel_com_cnpj['NF_PAINEL_PURA']
        
        chaves_lancadas = set(painel_com_cnpj[painel_com_cnpj['NF_PAINEL_PURA'] != ""]['chave_p'].unique())
        cnpjs_no_painel = set(painel_com_cnpj['CNPJ_FORN_LIMPO'].unique())

        # 4. Cruzamento Pedidos (Correção do Erro de Coluna)
        df_relacao['Cód. fornecedor_Limpo'] = df_relacao['Cód. fornecedor'].apply(limpar_cod)
        rel_com_cnpj = pd.merge(
            df_relacao, 
            df_forn[['Cód. Fornecedor_Limpo', 'CNPJ_FORN_LIMPO']], 
            left_on='Cód. fornecedor_Limpo', 
            right_on='Cód. Fornecedor_Limpo', 
            how='left'
        )
        peds_agrupados = rel_com_cnpj.groupby('CNPJ_FORN_LIMPO')['Nº do pedido'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index()

        # 5. Cruzamento Contratos
        registros_ct = []
        item_atual = {'Contrato': None, 'CNPJ': None}
        for i in range(len(df_bruto_ct)):
            l = df_bruto_ct.iloc[i]
            col_a = str(l[0]).strip() if pd.notna(l[0]) else ""
            if col_a == "Contrato": item_atual['Contrato'] = str(l[3]).strip()
            elif col_a == "CNPJ" and item_atual['Contrato']:
                item_atual['CNPJ'] = limpar_cnpj(l[3])
                registros_ct.append(item_atual.copy())
        
        cts_agrupados = pd.DataFrame(registros_ct).groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index() if registros_ct else pd.DataFrame(columns=['CNPJ', 'Contrato'])

        # 6. Unificação e Status
        resumo = pd.merge(df_nf, painel_com_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), left_on='chave_unica', right_on='chave_p', how='left')
        resumo = pd.merge(resumo, peds_agrupados, left_on='CNPJ_EMIT_LIMPO', right_on='CNPJ_FORN_LIMPO', how='left')
        resumo = pd.merge(resumo, cts_agrupados, left_on='CNPJ_EMIT_LIMPO', right_on='CNPJ', how='left')

        def definir_status(r):
            if r['chave_unica'] in chaves_lancadas: return "✅ NF Lançada"
            if r['CNPJ_EMIT_LIMPO'] in set(peds_agrupados['CNPJ_FORN_LIMPO']): return "⚠️ Para Verificação"
            if pd.notna(r['Contrato']): return "📄 Vínculo Contratual"
            return "❌ Sem Histórico"

        resumo['Status'] = resumo.apply(definir_status, axis=1)

        # Exportação
        final_df = resumo[['Núm/Série', 'CNPJ emitente', 'Emitente', 'Emissão', 'Valor', 'N° da Nota fiscal', 'Nº do pedido', 'Contrato', 'Status']]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False)
        
        st.success("Tudo pronto! Relatório de Auditoria gerado com sucesso.")
        st.download_button("📥 Baixar Auditoria", output.getvalue(), "AUDITORIA_NF_PRODUTO.xlsx")
