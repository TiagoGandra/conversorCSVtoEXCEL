import streamlit as st
import pandas as pd
import io  # Necessário para trabalhar com bytes em memória
# XlsxWriter é usado pelo pandas como engine, não precisa importar diretamente aqui
# mas precisa estar instalado (pip install xlsxwriter)

# Título da Aplicação
st.title("Visualizador e Conversor de CSV para Excel")

# --- Barra Lateral ---
with st.sidebar:
    st.header("Desenvolvido com Streamlit e Pandas por Tiago Gandra :)")

    # 1. Upload do arquivo CSV na barra lateral
    uploaded_file = st.file_uploader("Carregue seu arquivo CSV aqui", type="csv")

# --- Área Principal ---

# Verifica se um arquivo foi carregado
if uploaded_file is not None:
    try:
        # Lê o arquivo CSV para um DataFrame do Pandas
        # Tenta detectar o separador comum (vírgula ou ponto e vírgula)
        try:
            df = pd.read_csv(uploaded_file, sep=',')
        except Exception: # Se falhar com vírgula, tenta ponto e vírgula
             # Garante que o ponteiro do arquivo volte ao início
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=';')

        st.success("Arquivo CSV carregado com sucesso!")

        # 2. Mostra o DataFrame como tabela
        st.header("Dados do Arquivo CSV")
        st.dataframe(df) # st.dataframe é interativo, st.table é estático

        # --- Preparação para o Download COM ESTILIZAÇÃO ---

        # Cria um buffer de bytes em memória para salvar o Excel
        output = io.BytesIO()

        # Use pd.ExcelWriter com o engine XlsxWriter
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Escreve o DataFrame (sem o índice do pandas)
            sheet_name = 'Dados_Estilizados'
            df.to_excel(writer, index=False, sheet_name=sheet_name)

            # Obtenha os objetos workbook e worksheet do XlsxWriter
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            # --- Definição dos Formatos ---
            # Formato para o cabeçalho (negrito, fundo cinza claro, centralizado)
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': False, # Evita quebra de linha no cabeçalho
                'valign': 'vcenter', # Alinhamento vertical centralizado
                'fg_color': '#D9D9D9', # Cinza claro
                'border': 1,
                'align': 'center' # Alinhamento horizontal centralizado
            })

            # --- Aplicação dos Formatos ---

            # Aplicar o formato do cabeçalho (Linha 1 do Excel = linha 0 aqui)
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Ajustar a largura das colunas automaticamente (método simples)
            for idx, col in enumerate(df.columns):
                series = df[col]
                # Calcula o comprimento máximo entre o nome da coluna e os dados da coluna
                # Adiciona 2 para um pequeno espaçamento extra
                max_len = max((
                    series.astype(str).map(len).max(), # Comprimento máx dos dados
                    len(str(series.name)) # Comprimento do nome da coluna
                )) + 2
                # Limita a largura máxima (opcional, para evitar colunas excessivamente largas)
                max_len = min(max_len, 50)
                worksheet.set_column(idx, idx, max_len) # Define a largura da coluna idx

        # Pega os bytes do buffer (após o 'with' terminar e salvar)
        excel_data = output.getvalue()

        # --- Botão de Download ---

        st.header("Exportar para Excel")

        # Gera um nome de arquivo dinâmico
        original_filename = uploaded_file.name
        if original_filename.lower().endswith('.csv'):
            download_filename = original_filename[:-4] + '.xlsx'
        else:
            download_filename = original_filename + '.xlsx'

        # Botão para baixar o arquivo .xlsx estilizado
        st.download_button(
            label="📥 Baixar como XLSX",
            data=excel_data, # Os dados em bytes do arquivo Excel estilizado
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        st.warning("Verifique se o arquivo é um CSV válido e se o separador é ',' ou ';'.")

else:
    st.info("Aguardando o upload de um arquivo CSV na barra lateral.")