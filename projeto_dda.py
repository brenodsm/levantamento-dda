import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# Carregar e preparar os dados
df_dda = pd.read_excel('Planejamento Pagamentos DDA via arquivo.xlsx')
df_dda = df_dda.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

st.set_page_config(layout="wide")


# Definir colunas e mapas de cores
coluna1_dda = 'Situação'
coluna2_dda = 'Nome Cedente'
coluna3_dda = 'Setor'
coluna4_dda = 'Observação do Vínculo'

color_map = {
    'Automático': 'green',
    'Manual': 'blue',
    'Pendente': 'yellow',
    'Vinculado - Com diferença': 'red'
}

color_map_vinculo = {
    'Número do documento diferente, Cnpj/Cpf do cedente': 'purple',
    'Número do documento diferente': 'red',
    'Cnpj/Cpf do cedente diferente': 'blue'
}

# Separar os setores por barra e expandir as linhas
df_dda_expanded = df_dda.assign(Setor=df_dda[coluna3_dda].str.split('/')).explode('Setor')
df_dda_expanded['Setor'] = df_dda_expanded['Setor'].str.strip()  # Remover espaços em branco

# Criar gráficos
grafico_situacao = px.pie(df_dda, names=coluna1_dda, title='Situação Vinculo', 
                          color=coluna1_dda, color_discrete_map=color_map)
grafico_situacao.update_traces(
    hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<extra></extra>'
)

tabela_agrupada = df_dda.groupby([coluna2_dda, coluna1_dda]).size().reset_index(name='count')
grafico_cruzado = px.bar(tabela_agrupada, x=coluna2_dda, y='count',
                         color=coluna1_dda, title='Distribuição de Situação por Nome Cedente',
                         color_discrete_map=color_map)

tabela_filtrada = df_dda[df_dda[coluna1_dda] == 'Vinculado - Com diferença']
tabela_filtrada_agrupada = tabela_filtrada.groupby([coluna2_dda, coluna4_dda]).size().reset_index(name='count')
grafico_vinculado = px.bar(tabela_filtrada_agrupada, x=coluna2_dda, y='count', 
                           color=coluna4_dda, title='Diferença por Nome Cedente',
                           color_discrete_map=color_map_vinculo)
grafico_pizza = px.pie(tabela_filtrada_agrupada, names=coluna4_dda, values='count',
                       title='Distribuição de Vinculado - Com Diferença por Observação do Vínculo',
                       color=coluna4_dda, color_discrete_map=color_map_vinculo)



# Gráfico por setor (com setores separados por barra)
tabela_setor_agrupada = df_dda_expanded.groupby([coluna3_dda, coluna1_dda]).size().reset_index(name='count')

# Adiciona a coluna 'Situação' como customdata para o gráfico
grafico_setor = px.bar(tabela_setor_agrupada, x=coluna3_dda, y='count',
                      color=coluna1_dda, title='Distribuição de Vínculo por Setor',
                      color_discrete_map=color_map, custom_data=[coluna1_dda])

# Atualiza o hovertemplate para mostrar Setor, Situação e Quantidade
grafico_setor.update_traces(
    hovertemplate='<b>Setor: %{x}</b><br>Situação: %{customdata[0]}<br>Quantidade: %{y}<extra></extra>'
)


# Criar a visualização
st.title('Levantamento DDA')
st.sidebar.title("**Selecione uma opção**")
opcoes_sidebar = st.sidebar.radio('ㅤ', ['Situação Vinculo', 'Situação - Nome Cedente', 'Diferença pelo nome das empresas', 'Vínculos Agrupados por Setor','Empresas por Setores'])


for c in range(10):
    st.sidebar.write('')

# Adicionando a imagem na barra lateral
st.sidebar.image('pzt_trans.png', caption='', use_column_width=True)



if opcoes_sidebar == 'Situação Vinculo':
    st.plotly_chart(grafico_situacao)

    st.subheader("Observações")
    
    with st.expander("30,8% Vinculado com Diferença"):
        st.markdown(f"""
        <span style="color:{color_map['Vinculado - Com diferença']}">Consegue fazer o vínculo manual e processar a remessa.  
        **Impacto Negativo:** O retrabalho na validação das informações leva o dobro do tempo em comparação ao método atual.</span>
        """, unsafe_allow_html=True)
    
    with st.expander("54% Pendente"):
        st.markdown(f"""
    <span>**Agrupamentos:** Erros com diferença de valor e lançamento posterior ao fechamento (até 7 dias). Requer o mesmo retrabalho do "Vinculado com Diferença", pois é necessário acessar outra tela.  
    Para ajustar a diferença de valor, é preciso realizar um ajuste manual.  
    **Vencimentos Errados:** Manutenção necessária.  
    **Ação Requerida:** Trazer por setor os vencimentos incorretos.</span>
    """, unsafe_allow_html=True)

    
    with st.expander("5,02% Manual"):
        st.markdown(f"""
        <span style="color:{color_map['Manual']}">Não consegue vincular automaticamente, resultando no mesmo retrabalho do "Vinculado com Diferença".  
        **Ação Requerida:** Buscar entendimento, pois a causa não foi identificada. Trazer evidências.</span>
        """, unsafe_allow_html=True)
    
    with st.expander("10,2% Automático"):
        st.markdown(f"""
        <span style="color:{color_map['Automático']}">É vinculado automaticamente, sem necessidade de intervenção manual.</span>
        """, unsafe_allow_html=True)

elif opcoes_sidebar == 'Situação - Nome Cedente':
    pesquisa = st.text_input('Pesquisar por Nome Cedente (separados por vírgula):', '')
    
    if pesquisa:
        termos_pesquisa = [term.strip() for term in pesquisa.split(',')]
        tabela_agrupada = tabela_agrupada[tabela_agrupada[coluna2_dda].apply(lambda x: any(term.lower() in x.lower() for term in termos_pesquisa))]
        grafico_cruzado = px.bar(tabela_agrupada, x=coluna2_dda, y='count',
                                 color=coluna1_dda, title='Distribuição de Situação por Nome Cedente',
                                 color_discrete_map=color_map)
    
    mostrar_nomes = st.checkbox('Mostrar nomes das empresas')
    
    if not mostrar_nomes:
        grafico_cruzado.update_xaxes(showticklabels=False)
    
    st.plotly_chart(grafico_cruzado)

elif opcoes_sidebar == 'Diferença pelo nome das empresas':
    pesquisa = st.text_input('Pesquisar por Nome Cedente (separados por vírgula):', '')
    
    if pesquisa:
        termos_pesquisa = [term.strip() for term in pesquisa.split(',')]
        tabela_filtrada_agrupada = tabela_filtrada_agrupada[tabela_filtrada_agrupada[coluna2_dda].apply(lambda x: any(term.lower() in x.lower() for term in termos_pesquisa))]
        grafico_vinculado = px.bar(tabela_filtrada_agrupada, x=coluna2_dda, y='count', 
                                   color=coluna4_dda, title='Diferença por Nome Cedente',
                                   color_discrete_map=color_map_vinculo)
    
    mostrar_nomes = st.checkbox('Mostrar nomes das empresas')
    
    if not mostrar_nomes:
        grafico_vinculado.update_xaxes(showticklabels=False)
    
    st.plotly_chart(grafico_vinculado)

elif opcoes_sidebar == 'Vínculos Agrupados por Setor':
    st.plotly_chart(grafico_setor)

 
elif opcoes_sidebar == 'Empresas por Setores':

    # Separar os setores divididos por barra '/' em várias linhas
    df_dda = df_dda.assign(Setor=df_dda['Setor'].str.split('/')).explode('Setor')

    # Remover duplicatas com base no 'Nome Cedente' e agrupar observações
    df_dda['Observação do Vínculo'] = df_dda.groupby(['Nome Cedente', 'Setor'])['Observação do Vínculo'].transform(lambda x: ', '.join(x.dropna().unique()))

    # Remover duplicatas, mantendo apenas uma linha por empresa e setor
    df_unicos = df_dda.drop_duplicates(subset=['Nome Cedente', 'Setor'])

    # Exibir o dropdown para selecionar o setor
    setores = df_unicos['Setor'].unique()
    setor_selecionado = st.selectbox("Escolha o Setor", setores)

    # Filtrar as empresas pelo setor selecionado
    empresas_filtradas = df_unicos[df_unicos['Setor'] == setor_selecionado]

    # Definir o mapeamento de cores para a coluna "Situação" com tons mais claros
    color_map = {
        'Automático': 'background-color: lightgreen',  # Verde claro
        'Manual': 'background-color: lightblue',       # Azul claro
        'Pendente': 'background-color: yellow',        # Mantém o amarelo
        'Vinculado - Com diferença': 'background-color: lightcoral'  # Vermelho claro
    }

    # Função para aplicar a coloração com base na "Situação"
    def apply_colors(val):
        return color_map.get(val, '')

    # Aplicar o estilo na coluna "Situação" para colorir as células
    styled_df = empresas_filtradas.style.applymap(apply_colors, subset=['Situação'])

    # Exibir a tabela com a formatação
    st.dataframe(styled_df)

    # Criar um buffer de memória para o arquivo Excel
    output = BytesIO()
    empresas_filtradas.to_excel(output, index=False)
    output.seek(0)  # Retorna ao início do buffer

    # Mostrar o botão de download
    st.download_button(
        label="Clique aqui para baixar o arquivo Excel",
        data=output,
        file_name="empresas_filtradas.xlsx",
        mime="application/vnd.ms-excel"
    )
