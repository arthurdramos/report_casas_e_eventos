from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import pandas as pd
import io
import streamlit as st
import plotly.express as px
from dotenv import load_dotenv
import os

# Carregar as variáveis do arquivo .env
load_dotenv("credentials.env")

# Obter os detalhes do SharePoint a partir das variáveis de ambiente
url_shrpt = os.getenv('URL_SHRPT')
username_shrpt = os.getenv('USERNAME_SHRPT')
password_shrpt = os.getenv('PASSWORD_SHRPT')
file_url_shrpt = os.getenv('FILE_URL_SHRPT')

# Autenticação
ctx_auth = AuthenticationContext(url_shrpt)
if ctx_auth.acquire_token_for_user(username_shrpt, password_shrpt):
    ctx = ClientContext(url_shrpt, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Autenticação bem-sucedida!")
else:
    print("Erro de autenticação:", ctx_auth.get_last_error())
    exit()

# Acessar o arquivo no SharePoint
response = File.open_binary(ctx, file_url_shrpt)

# Carregar o conteúdo do arquivo em um objeto BytesIO
bytes_file = io.BytesIO(response.content)

# Ler a aba "GeralEventos" da planilha
df = pd.read_excel(bytes_file, sheet_name="GeralCasas")

# Redefinir o índice para remover a coluna de índice original
df.reset_index(drop=True, inplace=True)

# Excluir a linha onde o "Comercial" é "Total Geral"
df = df[df['Cod Comercial'] != 'Total Geral']

# Certificar-se de que a coluna '%' é numérica e multiplicar por 100 para converter em porcentagem
df['%'] = pd.to_numeric(df['%'], errors='coerce') * 100

# Ordenar o dataframe pela coluna 'Pontuação Total' em ordem decrescente
df = df.sort_values(by='Soma de Pontuação Total', ascending=False)

# Adicionar CSS para centralizar o conteúdo
st.markdown(
    """
    <style>
    .centered-content {
        display: flex;
        justify-content: center;
        align-items: center;
        flex-direction: column;
    }
    </style>
    """, unsafe_allow_html=True
)

# Centralizar o conteúdo com uma div
st.markdown('<div class="centered-content">', unsafe_allow_html=True)

st.title('Relatório de Produtos:')
st.subheader('Pontuação Geral Casas')

# Exibir o dataframe completo primeiro, centralizado
st.title('Dados completos')
st.dataframe(df)

# Criar um filtro para selecionar o comercial
comercial_selecionado = st.selectbox('Selecione o Comercial', df['Cod Comercial'].unique())

# Filtrar o dataframe para o comercial selecionado (somente para o gráfico)
df_filtrado = df[df['Cod Comercial'] == comercial_selecionado]

# Criar o gráfico usando Plotly, com rótulos de valor nas barras
fig = px.bar(df_filtrado, x='Cod Comercial', y=['Soma de Pontuação Total', 'Meta', '%'],
             title=f'Dados do Comercial: {comercial_selecionado}',
             labels={'value': 'Valores', 'variable': 'Indicadores'},
             barmode='group',
             text_auto=True)

# Atualizar o trace apenas para a coluna '%', formatando-a como porcentagem
fig.for_each_trace(lambda t: t.update(texttemplate='%{y:.2f}%', textposition='outside') if t.name == '%' else t)

# Exibir o gráfico no Streamlit
st.plotly_chart(fig)

# Fechar a div centralizada
st.markdown('</div>', unsafe_allow_html=True)
