import pandas as pd
import streamlit as st
from io import BytesIO

# Função para processar o arquivo Excel
def processar_arquivo(uploaded_file, file_type):
    try:
        if file_type == 'xlsb':
            df = pd.read_excel(uploaded_file, engine='pyxlsb')
        elif file_type == 'xlsx':
            df = pd.read_excel(uploaded_file)
        elif file_type == 'csv':
            df = pd.read_csv(uploaded_file)
        else:
            st.error("Tipo de arquivo não suportado!")
            return None

        # Verifica se as colunas necessárias estão presentes
        colunas_necessarias = ['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Glosa Inicial', 'Desc. Proced. DCM']
        if not all(col in df.columns for col in colunas_necessarias):
            st.error("O arquivo não contém as colunas necessárias!")
            return None

        # Selecionando as colunas necessárias
        colunas = ['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Glosa Inicial', 'Desc. Proced. DCM']
        df_selecionado = df[colunas]

        # Agrupando por 'Motivo Operadora (Código)' e 'Operadora' e somando os valores
        df_somatorio = df_selecionado.groupby(['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Desc. Proced. DCM'])[
            ['Glosa Inicial']
        ].sum().reset_index()

        # Ordenando o DataFrame pelo 'Motivo Operadora (Código)' e 'Operadora'
        df_somatorio = df_somatorio.sort_values(by=['Marca', 'Operadora', 'Motivo Operadora (Código)'])

        # Função para formatar o valor em uma string legível
        def formatar_valor(valor):
            if valor >= 1e6:
                return f"{valor/1e6:.2f}M"
            elif valor >= 1e3:
                return f"{valor/1e3:.2f}k"
            else:
                return f"{valor:.0f}"

        # Função para formatar a linha como uma string, incluindo a numeração
        def formatar_linha(row, contador):
            valor_formatado = formatar_valor(row['Glosa Inicial'])
            return f"{contador} - {row['Desc. Proced. DCM']} ({valor_formatado})"

        def linha_completa(dados):
            # Reiniciar o índice para numerar corretamente as linhas
            dados.reset_index(drop=True, inplace=True)

            # Aplicar a função formatar_linha a cada linha e criar uma nova coluna 'String Formatada'
            dados['String Formatada'] = dados.apply(lambda row: formatar_linha(row, row.name + 1), axis=1)

            # Juntar todas as strings em uma única string separada por quebra de linha
            todas_as_linhas = '\n'.join(dados['String Formatada'])
            return todas_as_linhas

        df_principal = pd.DataFrame(columns=['Unidade', 'Operadora', 'Código de Glosa', 'Valor Glosado', 'Ofensores (TOP 5)'])

        # Selecionando as colunas necessárias
        colunas = ['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Glosa Inicial']
        df_selecionado2 = df[colunas]

        # Agrupando por 'Motivo Operadora (Código)' e 'Operadora' e somando os valores
        df_somatorio_marca = df_selecionado2.groupby(['Marca', 'Operadora', 'Motivo Operadora (Código)'])[
            ['Glosa Inicial']
        ].sum().reset_index()

        df_soma_operadora = df_selecionado2.groupby(['Marca', 'Operadora'])[
            ['Glosa Inicial']
        ].sum().reset_index()

        # Ordenando o DataFrame pelo 'Motivo Operadora (Código)' e 'Operadora'
        df_somatorio_marca = df_somatorio_marca.sort_values(by=['Marca', 'Operadora', 'Motivo Operadora (Código)'])

        # Primeiro, obtenha as marcas únicas
        lista_marcas = df_somatorio_marca['Marca'].unique()

        # Em seguida, iteramos sobre cada marca
        for marca in lista_marcas:
            df_1marca = df_soma_operadora[df_soma_operadora['Marca'] == marca].sort_values('Glosa Inicial', ascending=False)

            df_marca_codigo = df_somatorio_marca[df_somatorio_marca['Marca'] == marca].sort_values('Glosa Inicial', ascending=False)

            lista_operadoras = df_1marca['Operadora'].unique()
            # Iterando sobre o operador de cada marca
            for operadora in lista_operadoras:
                df_1operadora = df_marca_codigo[df_marca_codigo['Operadora'] == operadora].sort_values('Glosa Inicial', ascending=False)[:5]

                codigos = df_1operadora['Motivo Operadora (Código)'].unique()

                for codigo_ in codigos:
                    df_top5_dcm_ = df_somatorio.query(f"Operadora == '{operadora}' and Marca == '{marca}' and `Motivo Operadora (Código)` =='{codigo_}' ").sort_values('Glosa Inicial', ascending=False).reset_index()[:5]
                    texto = linha_completa(df_top5_dcm_)
                    soma_total = df_somatorio_marca.query(f"Operadora == '{operadora}' and Marca == '{marca}' and `Motivo Operadora (Código)` =='{codigo_}'")['Glosa Inicial'].values[0]
                    dados = [marca, operadora, codigo_, soma_total, texto]

                    # Criando um novo DataFrame com os dados
                    novo_df = pd.DataFrame([dados], columns=df_principal.columns)

                    # Concatenando o DataFrame original com o novo DataFrame
                    df_principal = pd.concat([df_principal, novo_df], ignore_index=True)

        return df_principal
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None

# Função para converter o DataFrame para o formato Excel em um buffer
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

# Função principal do Streamlit
def main():
    st.title("Plano de Ação Glosa Inicial - Desenvolvido por Hebert Garcia")

    uploaded_file = st.file_uploader("Faça upload do arquivo Excel", type=["xlsb", "xlsx", "csv"])

    if uploaded_file is not None:
        file_type = uploaded_file.name.split('.')[-1]
        
        if st.button("Processar Arquivo"):
            df_resultado = processar_arquivo(uploaded_file, file_type)

            if df_resultado is not None:
                st.write("Arquivo processado com sucesso!")
                st.write(df_resultado)

                # Converter o DataFrame para o formato Excel
                df_xlsx = convert_df_to_excel(df_resultado)

                st.download_button(
                    label="Baixar arquivo Excel formatado",
                    data=df_xlsx,
                    file_name="resultado_formatado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
