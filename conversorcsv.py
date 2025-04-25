import streamlit as st
import pandas as pd
import io  # Necessário para trabalhar com bytes em memória

# Título da Aplicação
st.title("Visualizador e Conversor de CSV para Excel")

with st.sidebar:
    st.header("Desenvolvido com Streamlit e Pandas por Tiago Gandra :)")

    # 1. Upload do arquivo CSV
    uploaded_file = st.file_uploader("Carregue seu arquivo CSV aqui", type="csv")

# Verifica se um arquivo foi carregado
if uploaded_file is not None:
    try:
        # Lê o arquivo CSV para um DataFrame do Pandas
        df = pd.read_csv(uploaded_file)

        st.success("Arquivo CSV carregado com sucesso!")

        # 2. Mostra o DataFrame como tabela
        st.header("Dados do Arquivo CSV")
        st.dataframe(df) # st.dataframe é interativo, st.table é estático

        # --- Preparação para o Download ---

        # Cria um buffer de bytes em memória para salvar o Excel
        output = io.BytesIO()

        # Escreve o DataFrame no buffer em formato Excel (.xlsx)
        # 'index=False' evita que o índice do DataFrame seja escrito no Excel
        # 'engine='openpyxl'' é necessário para o formato .xlsx
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dados')
            # Você pode adicionar mais planilhas se necessário aqui
            # df_outra_coisa.to_excel(writer, index=False, sheet_name='OutrosDados')

        # Pega os bytes do buffer
        excel_data = output.getvalue()

        # --- Botão de Download ---

        st.header("Exportar para Excel")

        # Gera um nome de arquivo dinâmico (opcional, mas útil)
        # Pega o nome do arquivo original e troca a extensão
        original_filename = uploaded_file.name
        if original_filename.lower().endswith('.csv'):
            download_filename = original_filename[:-4] + '.xlsx'
        else:
            download_filename = original_filename + '.xlsx'


        # 3. Botão para baixar o arquivo .xlsx
        st.download_button(
            label="📥 Baixar como XLSX",
            data=excel_data, # Os dados em bytes do arquivo Excel
            file_name=download_filename, # O nome que o arquivo terá ao ser baixado
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" # O tipo MIME para arquivos .xlsx
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        st.warning("Verifique se o arquivo é um CSV válido.")

else:
    st.info("Aguardando o upload de um arquivo CSV.")
