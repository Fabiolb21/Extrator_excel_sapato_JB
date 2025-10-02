import streamlit as st
import pandas as pd
import numpy as np
import os
import zipfile
import xlwt
from io import BytesIO

st.set_page_config(page_title="Processador de Etiquetas", layout="wide")

st.title("📊 Processador de Planilhas para Etiquetas")

st.markdown("""
Esta aplicação processa uma planilha Excel para extrair e transformar dados, gerando novas planilhas prontas para a criação de etiquetas.

**Instruções:**
1. Faça o upload do seu arquivo Excel (formato `.xlsm` ou `.xlsx`).
2. Aguarde o processamento dos dados.
3. Faça o download do arquivo ZIP contendo as planilhas geradas.
""")

st.divider()


# Função para processar o arquivo
def process_excel(uploaded_file):
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(uploaded_file, sheet_name='Dados', engine='openpyxl')
        st.success(f"✓ Arquivo Excel lido com sucesso! ({len(df)} linhas encontradas)")

        # Criar as novas colunas
        st.info("🔄 Criando novas colunas...")
        df['PRODUTO'] = df['PROD_DESCRICAO'].apply(lambda x: str(x).split(' ')[0] if pd.notna(x) else '')
        df['DESCRICAO'] = df['PROD_DESCRICAO'].apply(lambda x: ' '.join(str(x).split(' ')[1:]) if pd.notna(x) else '')
        df['PROD_DESC'] = df['PROD_DESCRICAO'].str.slice(0, 9)
        df['IMAGEM_MODELO_NEW'] = df['PROD_DESC'].apply(
            lambda x: f"\\\\SERVER-DADOS\\Label\\CÓDIGOS\\H.Kuntzler\\JORGE BISCHOFF\\SAPATOS\\{x}.jpg" if pd.notna(x) else ''
        )
        
        # Formatar PRECO_UNIT_PDV com R$
        df['PRECO_UNIT_PDV'] = df['PRECO_UNIT_PDV'].apply(
            lambda x: f"R$ {float(x):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else ''
        )
        
        # Substituir NaN por string vazia na coluna LARGURA
        df['LARGURA'] = df['LARGURA'].apply(lambda x: '' if pd.isna(x) else str(x))

        # Selecionar e reordenar as colunas
        colunas_finais = [
            'PLANO_PROD', 'OF_NUMERO', 'PROD_DESCRICAO', 'PRODUTO', 'DESCRICAO',
            'PROD_CODIGO', 'PRECO_UNIT_PDV', 'LARGURA', 'GRADE_TAMANHO',
            'CODIGO_BARRAS', 'QTD', 'UNID_MEDIDA', 'PROD_DESC', 'IMAGEM_MODELO_NEW'
        ]
        df_final = df[colunas_finais]

        # Duplicar linhas com base na coluna QTD
        st.info("🔄 Duplicando linhas conforme a quantidade...")
        df_expandido = df_final.loc[df_final.index.repeat(df_final['QTD'])].reset_index(drop=True)
        st.success(f"✓ Total de linhas após duplicação: {len(df_expandido)}")

        # Gerar planilhas por OF_NUMERO
        st.info("🔄 Gerando planilhas individuais por OF_NUMERO...")
        output_dir = "output_files"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        grouped = df_expandido.groupby('OF_NUMERO')
        file_paths = []
        
        # Barra de progresso
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_groups = len(grouped)
        
        for idx, (of_numero, group) in enumerate(grouped):
            # Atualizar progresso
            progress = (idx + 1) / total_groups
            progress_bar.progress(progress)
            status_text.text(f"Processando OF_NUMERO {of_numero} ({idx + 1}/{total_groups})...")
            
            # Converter todas as colunas para string, mantendo vazios sem NaN
            group_copy = group.copy()
            for col in group_copy.columns:
                group_copy[col] = group_copy[col].apply(lambda x: '' if pd.isna(x) or str(x).lower() == 'nan' else str(x))

            # Criar arquivo .xls usando xlwt
            output_filename = os.path.join(output_dir, f"{of_numero}.xls")
            
            workbook = xlwt.Workbook()
            worksheet = workbook.add_sheet('Dados')

            # Escrever cabeçalhos
            for col_idx, col_name in enumerate(group_copy.columns):
                worksheet.write(0, col_idx, col_name)

            # Escrever dados
            for row_idx, row in enumerate(group_copy.values, start=1):
                for col_idx, value in enumerate(row):
                    # Garantir que não escreve 'nan' ou 'NaN'
                    val_str = str(value) if value != '' and str(value).lower() != 'nan' else ''
                    worksheet.write(row_idx, col_idx, val_str)

            workbook.save(output_filename)
            file_paths.append(output_filename)

        progress_bar.empty()
        status_text.empty()
        st.success(f"✓ {len(file_paths)} planilhas geradas com sucesso!")
        
        return file_paths, df_expandido

    except Exception as e:
        st.error(f"❌ Ocorreu um erro: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None, None


# Interface do Streamlit
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader("📁 Escolha um arquivo Excel", type=["xlsm", "xlsx"])

with col2:
    st.markdown("### Informações")
    st.markdown("""
    **Formato aceito:** `.xlsm`, `.xlsx`
    
    **Aba necessária:** `Dados`
    """)

if uploaded_file is not None:
    st.divider()
    
    if st.button("🚀 Processar Arquivo", type="primary", use_container_width=True):
        with st.spinner('Processando...'):
            file_paths, df_expandido = process_excel(uploaded_file)

            if file_paths:
                st.divider()
                st.subheader("📦 Download dos Resultados")
                
                # Mostrar estatísticas
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total de Planilhas", len(file_paths))
                with col2:
                    st.metric("Total de Linhas", len(df_expandido))
                with col3:
                    st.metric("OF_NUMERO Únicos", df_expandido['OF_NUMERO'].nunique())
                
                # Criar um arquivo ZIP com todas as planilhas
                zip_filename = "planilhas_geradas.zip"
                with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for file in file_paths:
                        zipf.write(file, os.path.basename(file))

                with open(zip_filename, "rb") as fp:
                    st.download_button(
                        label="⬇️ Download Planilhas (ZIP)",
                        data=fp,
                        file_name=zip_filename,
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
                
                # Mostrar preview dos dados
                with st.expander("👁️ Visualizar Preview dos Dados Processados"):
                    st.dataframe(df_expandido.head(50), use_container_width=True)

st.divider()
st.markdown("---")
st.markdown("**Desenvolvido com Streamlit** | Processador de Etiquetas v1.1")
