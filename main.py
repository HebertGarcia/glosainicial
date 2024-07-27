import pandas as pd
import streamlit as st
from io import BytesIO

def processar_arquivo_glosa_inicial(uploaded_file, file_type):
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

        colunas_necessarias = ['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Glosa Inicial', 'Desc. Proced. DCM']
        if not all(col in df.columns for col in colunas_necessarias):
            st.error("O arquivo não contém as colunas necessárias!")
            return None

        colunas = ['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Glosa Inicial', 'Desc. Proced. DCM']
        df_selecionado = df[colunas]

        df_somatorio = df_selecionado.groupby(['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Desc. Proced. DCM'])[
            ['Glosa Inicial']
        ].sum().reset_index()

        df_somatorio = df_somatorio.sort_values(by=['Marca', 'Operadora', 'Motivo Operadora (Código)'])

        def formatar_valor(valor):
            if valor >= 1e6:
                return f"{valor/1e6:.2f}M"
            elif valor >= 1e3:
                return f"{valor/1e3:.2f}k"
            else:
                return f"{valor:.0f}"

        def formatar_linha(row, contador):
            valor_formatado = formatar_valor(row['Glosa Inicial'])
            return f"{contador} - {row['Desc. Proced. DCM']} ({valor_formatado})"

        def linha_completa(dados):
            dados.reset_index(drop=True, inplace=True)
            dados['String Formatada'] = dados.apply(lambda row: formatar_linha(row, row.name + 1), axis=1)
            todas_as_linhas = '\n'.join(dados['String Formatada'])
            return todas_as_linhas

        df_principal = pd.DataFrame(columns=['Unidade', 'Operadora', 'Código de Glosa', 'Valor Glosado', 'Ofensores (TOP 5)'])

        colunas = ['Marca', 'Operadora', 'Motivo Operadora (Código)', 'Glosa Inicial']
        df_selecionado2 = df[colunas]

        df_somatorio_marca = df_selecionado2.groupby(['Marca', 'Operadora', 'Motivo Operadora (Código)'])[
            ['Glosa Inicial']
        ].sum().reset_index()

        df_soma_operadora = df_selecionado2.groupby(['Marca', 'Operadora'])[
            ['Glosa Inicial']
        ].sum().reset_index()

        df_somatorio_marca = df_somatorio_marca.sort_values(by=['Marca', 'Operadora', 'Motivo Operadora (Código)'])

        lista_marcas = df_somatorio_marca['Marca'].unique()

        for marca in lista_marcas:
            df_1marca = df_soma_operadora[df_soma_operadora['Marca'] == marca].sort_values('Glosa Inicial', ascending=False)
            df_marca_codigo = df_somatorio_marca[df_somatorio_marca['Marca'] == marca].sort_values('Glosa Inicial', ascending=False)
            lista_operadoras = df_1marca['Operadora'].unique()
            for operadora in lista_operadoras:
                df_1operadora = df_marca_codigo[df_marca_codigo['Operadora'] == operadora].sort_values('Glosa Inicial', ascending=False)[:8]
                codigos = df_1operadora['Motivo Operadora (Código)'].unique()
                for codigo_ in codigos:
                    df_top5_dcm_ = df_somatorio.query(f"Operadora == '{operadora}' and Marca == '{marca}' and `Motivo Operadora (Código)` =='{codigo_}' ").sort_values('Glosa Inicial', ascending=False).reset_index()[:5]
                    texto = linha_completa(df_top5_dcm_)
                    soma_total = df_somatorio_marca.query(f"Operadora == '{operadora}' and Marca == '{marca}' and `Motivo Operadora (Código)` =='{codigo_}'")['Glosa Inicial'].values[0]
                    dados = [marca, operadora, codigo_, soma_total, texto]
                    novo_df = pd.DataFrame([dados], columns=df_principal.columns)
                    df_principal = pd.concat([df_principal, novo_df], ignore_index=True)

        return df_principal
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None

def processar_arquivo_glosa_aceita(uploaded_file, file_type):
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

        colunas_necessarias = ['Marca', 'Operadora', 'Motivo Glosa Operadora (Código)', 'Glosa Aceita', 'Procedimento (Descrição)']
        if not all(col in df.columns for col in colunas_necessarias):
            st.error("O arquivo não contém as colunas necessárias!")
            return None

        colunas = ['Marca', 'Operadora', 'Motivo Glosa Operadora (Código)', 'Glosa Aceita', 'Procedimento (Descrição)']
        df_selecionado = df[colunas]

        df_somatorio = df_selecionado.groupby(['Marca', 'Operadora', 'Motivo Glosa Operadora (Código)', 'Procedimento (Descrição)'])[
            ['Glosa Aceita']
        ].sum().reset_index()

        df_somatorio = df_somatorio.sort_values(by=['Marca', 'Operadora', 'Motivo Glosa Operadora (Código)'])

        def formatar_valor(valor):
            if valor >= 1e6:
                return f"{valor/1e6:.2f}M"
            elif valor >= 1e3:
                return f"{valor/1e3:.2f}k"
            else:
                return f"{valor:.0f}"

        def formatar_linha(row, contador):
            valor_formatado = formatar_valor(row['Glosa Aceita'])
            return f"{contador} - {row['Procedimento (Descrição)']} ({valor_formatado})"

        def linha_completa(dados):
            dados.reset_index(drop=True, inplace=True)
            dados['String Formatada'] = dados.apply(lambda row: formatar_linha(row, row.name + 1), axis=1)
            todas_as_linhas = '\n'.join(dados['String Formatada'])
            return todas_as_linhas

        df_principal = pd.DataFrame(columns=['Unidade', 'Operadora', 'Código de Glosa', 'Valor Glosado', 'Ofensores (TOP 8)'])

        colunas = ['Marca', 'Operadora', 'Motivo Glosa Operadora (Código)', 'Glosa Aceita']
        df_selecionado2 = df[colunas]

        df_somatorio_marca = df_selecionado2.groupby(['Marca', 'Operadora', 'Motivo Glosa Operadora (Código)'])[
            ['Glosa Aceita']
        ].sum().reset_index()

        df_soma_operadora = df_selecionado2.groupby(['Marca', 'Operadora'])[
            ['Glosa Aceita']
        ].sum().reset_index()

        df_somatorio_marca = df_somatorio_marca.sort_values(by=['Marca', 'Operadora', 'Motivo Glosa Operadora (Código)'])

        lista_marcas = df_somatorio_marca['Marca'].unique()

        for marca in lista_marcas:
            df_1marca = df_soma_operadora[df_soma_operadora['Marca'] == marca].sort_values('Glosa Aceita', ascending=False)
            df_marca_codigo = df_somatorio_marca[df_somatorio_marca['Marca'] == marca].sort_values('Glosa Aceita', ascending=False)
            lista_operadoras = df_1marca['Operadora'].unique()
            for operadora in lista_operadoras:
                df_1operadora = df_marca_codigo[df_marca_codigo['Operadora'] == operadora].sort_values('Glosa Aceita', ascending=False)[:8]
                codigos = df_1operadora['Motivo Glosa Operadora (Código)'].unique()
                for codigo_ in codigos:
                    df_top5_dcm_ = df_somatorio.query(f"Operadora == '{operadora}' and Marca == '{marca}' and `Motivo Glosa Operadora (Código)` =='{codigo_}' ").sort_values('Glosa Aceita', ascending=False).reset_index()[:5]
                    texto = linha_completa(df_top5_dcm_)
                    soma_total = df_somatorio_marca.query(f"Operadora == '{operadora}' and Marca == '{marca}' and `Motivo Glosa Operadora (Código)` =='{codigo_}'")['Glosa Aceita'].values[0]
                    dados = [marca, operadora, codigo_, soma_total, texto]
                    novo_df = pd.DataFrame([dados], columns=df_principal.columns)
                    df_principal = pd.concat([df_principal, novo_df], ignore_index=True)

        return df_principal
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("Plano de Ação Glosas - Desenvolvido por Hebert Garcia")

    st.header("Glosa Inicial")
    uploaded_file_inicial = st.file_uploader("Faça upload do arquivo Excel de Glosa Inicial", type=["xlsb", "xlsx", "csv"], key="glosa_inicial")
    if uploaded_file_inicial is not None:
        file_type_inicial = uploaded_file_inicial.name.split('.')[-1]
        if st.button("Processar Arquivo Glosa Inicial"):
            df_resultado_inicial = processar_arquivo_glosa_inicial(uploaded_file_inicial, file_type_inicial)
            if df_resultado_inicial is not None:
                st.write("Arquivo de Glosa Inicial processado com sucesso!")
                st.write(df_resultado_inicial)
                df_xlsx_inicial = convert_df_to_excel(df_resultado_inicial)
                st.download_button(
                    label="Baixar arquivo Excel formatado de Glosa Inicial",
                    data=df_xlsx_inicial,
                    file_name="resultado_formatado_inicial.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    st.header("Glosa Aceita")
    uploaded_file_aceita = st.file_uploader("Faça upload do arquivo Excel de Glosa Aceita", type=["xlsb", "xlsx", "csv"], key="glosa_aceita")
    if uploaded_file_aceita is not None:
        file_type_aceita = uploaded_file_aceita.name.split('.')[-1]
        if st.button("Processar Arquivo Glosa Aceita"):
            df_resultado_aceita = processar_arquivo_glosa_aceita(uploaded_file_aceita, file_type_aceita)
            if df_resultado_aceita is not None:
                st.write("Arquivo de Glosa Aceita processado com sucesso!")
                st.write(df_resultado_aceita)
                df_xlsx_aceita = convert_df_to_excel(df_resultado_aceita)
                st.download_button(
                    label="Baixar arquivo Excel formatado de Glosa Aceita",
                    data=df_xlsx_aceita,
                    file_name="resultado_formatado_aceita.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
