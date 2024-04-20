"""
Projeto: Farol de Vendas - Dashboard Interativo

* @copyrigth    Sávio Silas <svosilas@gmail.com> - DEV Portal Vidros                
* @date         17 Abril 2024
* @file         main.py

* @brief
    Este script é responsável por criar um dashboard interativo de vendas utilizando Dash,
    uma framework Python para construção de aplicações web.

    O dashboard oferece as seguintes funcionalidades:
    - Filtragem de dados por vendedor.
    - Visualização da meta geral e realização em tempo real.
    - Projeções de vendas com base nos dados atuais.
    - Pontuação destacada do vendedor com base no desempenho.
    - Gráficos interativos de vendas por categoria e localidade.
    - Tabelas de faturamento detalhadas por subcategorias de produtos.
    Callbacks são utilizados para atualizar os componentes do aplicativo em resposta às interações do usuário.
    A estrutura do layout é organizada utilizando dash_bootstrap_components
    O dashboard é iniciado em um servidor local e pode ser acessado via navegador web para
    uma interação ao vivo com os dados.
"""
import dash,base64
import dash_bootstrap_components as dbc
from dash import html, dcc, Input, Output, State, dcc, dash_table
import pandas as pd
import mysql.connector
import numpy as np
import plotly.graph_objs as go
from datetime import datetime, timedelta
from io import BytesIO
from dash import dash_table
from dash_table.Format import Format, Scheme, Symbol, Group
from dash.dependencies import Input, Output, State, MATCH, ALL
from pandas.tseries.offsets import MonthEnd, BDay
import dash_auth
from flask import request
import os
import locale
import platform

if platform.system() == 'Windows':
    locale.setlocale(locale.LC_TIME, 'portuguese_brazil')
else:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Consultas
def fetch_data():
    config = {
        'user': 's',
        'password': 'a',
        'host': 'v',
        'database': 'i'
    }

    conn = mysql.connector.connect(**config)
    query = '''
    SELECT 
        iavos
    '''
    df = pd.read_sql(query, conn)
    conn.close()
    return df
df = fetch_data()
df['Vendedor'] = df['Vendedor'].apply(lambda x: x.split()[0])
df_metas = pd.read_excel("META_VENDEDORES.xlsx")
meta_geral_valor = df_metas['META GERAL'].values[0]
meta_geral_formatada = f"R$ {meta_geral_valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def fetch_data_benef():
    config = {
        'user': 'i',
        'password': 's',
        'host': 'a',
        'database': 'l'
    }

    conn = mysql.connector.connect(**config)
    query = '''
    lasis
    '''
    df_benef = pd.read_sql(query, conn)
    conn.close()
    return df_benef

def fetch_data_frete():
    config = {
        'user': 'a',
        'password': 'v',
        'host': 'i',
        'database': 'o'
    }

    conn = mysql.connector.connect(**config)
    query = '''
    SELECT
        avios
    '''
    df_frete = pd.read_sql(query, conn)
    conn.close()
    return df_frete

VALID_USERNAME_PASSWORD_PAIRS = {}
nomes_meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
auth = dash_auth.BasicAuth(app, VALID_USERNAME_PASSWORD_PAIRS)
app.server.secret_key = ''
app.server.secret_key = os.environ.get('', '')

@app.callback(
    Output('vendedor-dropdown', 'value'),
    [Input('interval-update', 'n_intervals'),
     Input('vendedor-dropdown', 'value')],
    [State('vendedor-dropdown', 'value')]
)
def update_vendedor_selecionado(n_intervals, selected_value, state_value):
    username = request.authorization['username']
    if username != '':
        return username
    return state_value


current_date = pd.to_datetime('today').normalize()
start_date_realizado = current_date.replace(day=1)
end_date_realizado = current_date

def calc_realizado(df_):
    ano_atual = datetime.now().year
    mes_atual = datetime.now().month
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'])
    df_filtered = df_[(df_['Data_Pedido'].dt.year == ano_atual) & (df_['Data_Pedido'].dt.month == mes_atual)]
    df_filtered_unique = df_filtered.drop_duplicates(subset='Id_Pedido', keep='first')

    return df_filtered_unique['TOTAL'].sum()

def calc_realizado_ate_ontem(df_):
    ano_atual = datetime.now().year
    mes_atual = datetime.now().month
    ontem = datetime.now().day-1

    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'])
    df_filtered = df_[(df_['Data_Pedido'].dt.year == ano_atual) &
                      (df_['Data_Pedido'].dt.month == mes_atual) &
                      (df_['Data_Pedido'].dt.day <= ontem)]
    df_filtered_unique = df_filtered.drop_duplicates(subset='Id_Pedido', keep='first')

    return df_filtered_unique['TOTAL'].sum()

def get_vendedor_names(df):
    first_names = [name.split()[0] for name in df['Vendedor'].unique()]
    options = [{'label': name, 'value': name} for name in first_names]
    options.insert(0, {'label': 'TODOS OS VENDEDORES', 'value': 'TODOS OS VENDEDORES'})
    return options

def dias_uteis_ate_ontem(start_date, end_date):
    # Converte as datas para o formato 'datetime64[D]'
    start_date_fmt = np.datetime64(start_date, 'D')
    # Ajusta end_date para incluir 'ontem' na contagem, adicionando um dia
    end_date_ajustado = end_date + np.timedelta64(1, 'D')
    end_date_fmt = np.datetime64(end_date_ajustado, 'D')
    
    return np.busday_count(start_date_fmt, end_date_fmt)

def total_dias_uteis_no_mes(start_date_realizado):
    start_date_realizado = start_date_realizado.replace(day=1)
    start_date_fmt = np.datetime64(start_date_realizado, 'D')
    # Calcula o último dia do mês
    last_day_of_month = start_date_realizado + pd.offsets.MonthEnd(1)
    last_day_of_month_fmt = np.datetime64(last_day_of_month, 'D')

    return np.busday_count(start_date_fmt, last_day_of_month_fmt + np.timedelta64(1, 'D'))

def calc_projecao_geral(valor_realizado):
    dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
    total_dias = total_dias_uteis_no_mes(start_date_realizado)
    if dias_uteis_ate_ontem_ > 0:  
        return (valor_realizado / dias_uteis_ate_ontem_) * total_dias
    return 0

@app.callback(
    Output("realizado-vendedor", "children"),
    [Input("vendedor-dropdown", "value")]
)
def calcular_realizado_vendedor(vendedor_selecionado):
    if vendedor_selecionado == 'TODOS OS VENDEDORES':
        return "R$ 0.00"
    
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y', errors='coerce')

    df_filtrado = df[(df['Vendedor'] == vendedor_selecionado) & 
                     (df['Data_Pedido'] >= start_date_realizado) & 
                     (df['Data_Pedido'] <= end_date_realizado)]
    df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')

    if df_filtrado_unico.empty:
        return "R$ 0.00"
    
    valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()
    valor_formatado = "R$ {:,.2f}".format(valor_realizado_vendedor).replace(",", "X").replace(".", ",").replace("X", ".")

    return valor_formatado

def calc_projecao_vendedor(df_, vendedor_selecionado):
    ano_atual = datetime.now().year
    mes_atual = datetime.now().month
    ontem = datetime.now().day-1
    # Garantindo que a coluna 'Data_Pedido' esteja no formato correto
    df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')

    # Filtrando o DataFrame pelo vendedor selecionado e pelo período
    df_filtrado = df_[(df_['Vendedor'] == vendedor_selecionado) & (df_['Data_Pedido'].dt.month == mes_atual) &
                      (df_['Data_Pedido'].dt.year == ano_atual)
                      & (df_['Data_Pedido'].dt.day <= ontem)]

    df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')
    valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()
    dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
    total_dias_uteis = total_dias_uteis_no_mes(start_date_realizado)

    # Calculando a projeção
    if dias_uteis_ate_ontem_ > 0: 
        projecao = (valor_realizado_vendedor / dias_uteis_ate_ontem_) * total_dias_uteis
    else:
        projecao = 0.00

    return projecao

@app.callback(
    Output("projecao-vendedor", "children"),
    [Input("vendedor-dropdown", "value")]
)
def atualizar_projecao_vendedor(vendedor_selecionado):
    if vendedor_selecionado == 'TODOS OS VENDEDORES':
        # Se nenhum vendedor estiver selecionado, não há o que calcular
        return "R$ 0.00"

    projecao = calc_projecao_vendedor(df, vendedor_selecionado)
    # Formatar a projeção como moeda
    projecao_formatada = "R$ {:,.2f}".format(projecao).replace(",", "X").replace(".", ",").replace("X", ".")

    return projecao_formatada

#################### GRÁFICO DE LINHA
def aggregate_daily_sales(df_, vendedor_selecionado):
    end_date = pd.to_datetime('today').normalize()
    start_date = end_date.replace(day=1)
    
    df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')
    
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df_filtered = df_[(df_['Data_Pedido'].between(start_date, end_date)) & (df_['Vendedor'] == vendedor_selecionado)]
    else:
        df_filtered = df_[df_['Data_Pedido'].between(start_date, end_date)]
        
    df_filtered_unique = df_filtered.drop_duplicates(subset='Id_Pedido', keep='first')
    
    daily_sales = df_filtered_unique.groupby('Data_Pedido')['TOTAL'].sum().reset_index()
    
    return daily_sales

@app.callback(
    Output('right-chart', 'figure'), 
    [Input('vendedor-dropdown', 'value')]
)
def update_line_chart(vendedor_selecionado):
    filtered_data = aggregate_daily_sales(df, vendedor_selecionado)
    return generate_line_chart(filtered_data)

# Função para gerar o gráfico de linha com os dados agregados
def generate_line_chart(agg_df):
    trace = go.Scatter(
        x=agg_df['Data_Pedido'],
        y=agg_df['TOTAL'],
        mode='lines+markers',
        name='Faturamento',
        line=dict(shape='spline', smoothing=1.3, width=4),
        marker=dict(color='#3B508C', size=10, line=dict(color='white', width=2)),
        hovertemplate='Dia %{x}<br>R$ %{y:,.2f}<extra></extra>',
        fill='tozeroy',
        fillcolor='rgba(91, 132, 188, 0.2)',
    )

    layout = go.Layout(
        title='FATURAMENTO POR DIA',
        xaxis=dict(title='', tickformat='%d/%m/%Y'),  
        yaxis=dict(
            title='',
            showgrid=True,
            gridcolor='white',
            tickprefix='R$ ',  
            tickformat=',.' 
        ),
        hovermode='closest',
        plot_bgcolor="white",
    
        autosize=True,
        margin=go.layout.Margin(l=30, r=30, b=50, t=50)
    )

    fig = go.Figure(data=[trace], layout=layout)
    fig.update_layout(
        yaxis_tickformat='R$,.2f',  
        separators=',.'  
    )

    return fig

end_date_3_meses = pd.to_datetime('today').normalize()
start_date_3_meses = (end_date_3_meses - pd.DateOffset(months=2)).replace(day=1)

#################### GRÁFICO DE PILHA
def apply_discount(row):
    if row['Tipo_Desconto'] == 'Porcentagem':
        return max(0, row['total_produto'] - (row['total_produto'] * row['Desconto'] / 100))
    elif row['Tipo_Desconto'] == 'Reais':
        valor_frete = row.get('Valor_Frete', 0)
        return ((row['total_produto'] * row['Desconto']) / ((row['TOTAL'] - valor_frete) + row['Desconto']) - row['total_produto']) * (-1)
    return row['total_produto'] 

def calcular_somas(df, categorias, inicio_mes, fim_mes, vendedor_selecionado):
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'])
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df_filtrado = df[(df['Data_Pedido'] >= inicio_mes) & 
                         (df['Data_Pedido'] <= fim_mes) & 
                         (df['Grupo'].isin(categorias)) & 
                         (df['Vendedor'] == vendedor_selecionado)]
    else:
        df_filtrado = df[(df['Data_Pedido'] >= inicio_mes) & 
                         (df['Data_Pedido'] <= fim_mes) & 
                         (df['Grupo'].isin(categorias))]

    df_filtrado['total_produto_com_desconto'] = df_filtrado.apply(apply_discount, axis=1)
    return df_filtrado['total_produto_com_desconto'].sum()
categorias_agregadas = ['ACESSÓRIOS', 'ALUMÍNIO', 'FERRAGEM', 'KIT PARA BOX PADRÃO', 'SILICONE']
categoria_vidro = ['VIDRO']
hoje = datetime.now()
meses = []
for i in range(2, 5):  
    inicio_mes = (hoje - pd.offsets.MonthBegin(n=i)).to_pydatetime()
    fim_mes = (inicio_mes + pd.offsets.MonthEnd(n=0)).to_pydatetime()
    meses.append((inicio_mes, fim_mes))

#################### Card Venda por Localidade 
def calcular_vendas_por_localidade(df, vendedor_selecionado=None):
    df = df.drop_duplicates(subset='Id_Pedido', keep='first')
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')
    
    mes_atual = datetime.now().month
    ano_atual = datetime.now().year
    df_mes_atual = df[(df['Data_Pedido'].dt.month == mes_atual) & (df['Data_Pedido'].dt.year == ano_atual)]
    
    if vendedor_selecionado and vendedor_selecionado != "TODOS OS VENDEDORES":
        df_mes_atual = df_mes_atual[df_mes_atual['Vendedor'] == vendedor_selecionado]
    
    vendas_capital = df_mes_atual[df_mes_atual['Cidade'].str.lower() == 'manaus']['TOTAL'].sum()
    vendas_interior = df_mes_atual[df_mes_atual['Cidade'].str.lower() != 'manaus']['TOTAL'].sum()
    
    return vendas_capital, vendas_interior

@app.callback(
    [Output('vendas-capital', 'children'),
     Output('vendas-interior', 'children')],
    [Input('vendedor-dropdown', 'value')]
)
def update_vendas_por_localidade(vendedor_selecionado):
    vendas_capital, vendas_interior = calcular_vendas_por_localidade(df, vendedor_selecionado)
    return [f"R$ {vendas_capital:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
            f"R$ {vendas_interior:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')]

# Callback para atualizar o card de "Clientes Atendidos - Grupos Específicos"
@app.callback(
    [Output('clientes-atendidos-vidro', 'children'),
     Output('clientes-atendidos-agregados', 'children'),
     Output('clientes-atendidos-temperados', 'children')],
    [Input('vendedor-dropdown', 'value')]
)
def update_clientes_atendidos(vendedor_selecionado):
    clientes_vidro, clientes_agregados, clientes_temperado = contar_clientes_grupos(df, vendedor_selecionado)
    return [
        f"QTD. {clientes_vidro}",
        f"QTD. {clientes_agregados}",
        f"QTD. {clientes_temperado}"
    ]

# #################### Card qtd clientes atendidos
def contar_clientes_grupos(df, vendedor_selecionado=None):
    grupos_agregados = ['ACESSÓRIOS', 'ALUMÍNIO', 'FERRAGEM', 'KIT PARA BOX PADRÃO', 'SILICONE']

    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')

    # Filtrar o DataFrame pelo mês e ano atual
    mes_atual = datetime.now().month
    ano_atual = datetime.now().year

    if vendedor_selecionado and vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    df_mes_atual = df[(df['Data_Pedido'].dt.month == mes_atual) & (df['Data_Pedido'].dt.year == ano_atual)]

    # Filtrar por grupos 'Agregados'
    df_agregados = df_mes_atual[df_mes_atual['Grupo'].isin(grupos_agregados)]
    clientes_agregados = df_agregados['Matriz_Cliente'].nunique()

    # Filtrar por grupo 'VIDRO' na matriz (considerado 'Temperado')
    df_temperado = df_mes_atual[(df_mes_atual['Grupo'] == 'VIDRO') & (df_mes_atual['Loja'] == 'PORTAL VIDROS (MATRIZ INDÚSTRIA)')]
    clientes_temperado = df_temperado['Matriz_Cliente'].nunique()

    # Filtrar por grupo 'VIDRO' na filial (considerado 'Vidro Comum'), excluindo 'TÁBUA' do subgrupo
    df_vidro_comum = df_mes_atual[(df_mes_atual['Grupo'] == 'VIDRO') & (df_mes_atual['Subgrupo'] != 'TÁBUA') & (df_mes_atual['Loja'] == 'PORTAL VIDROS (FILIAL)')]
    clientes_vidro_comum = df_vidro_comum['Matriz_Cliente'].nunique()

    return clientes_agregados, clientes_vidro_comum, clientes_temperado

# #################### Card RECOMPRA
@app.callback(
    Output("recompra-ultimos-6-meses", "children"),
    [Input("vendedor-dropdown", "value")]
)
def update_recompra_ultimos_6_meses(vendedor_selecionado):
    global df  # Indica que df é uma variável global
    end_date = pd.to_datetime("today").normalize()
    start_date = (end_date - pd.DateOffset(months=6)).replace(day=1)

    # Usa uma nova variável para o DataFrame filtrado
    df_filtrado = df
    if vendedor_selecionado and vendedor_selecionado != "TODOS OS VENDEDORES":
        df_filtrado = df[df['Vendedor'] == vendedor_selecionado]

    results = []

    for month_offset in range(7):  # Inclui um mês extra para cálculo de positivação
        month_start = start_date + pd.DateOffset(months=month_offset)
        month_end = (month_start + pd.DateOffset(months=1)) - pd.Timedelta(days=1)
        
        # Filtra o DataFrame já ajustado com base no vendedor
        df_month = df_filtrado[(df_filtrado['Data_Pedido'] >= month_start) & (df_filtrado['Data_Pedido'] <= month_end)]
        unique_clients = df_month['Matriz_Cliente'].nunique()
        results.append((month_start, unique_clients))

    children = []
  
      # Adicionando um título antes dos meses
    title = html.H5("RECOMPRA NOS ÚLTIMOS 6 MESES", className="card-title", style={"margin-bottom": "20px"})
    children.append(title)  # Adiciona o título à lista de children
    
    for i in reversed(range(1, 7)):
        current_month, current_count = results[i]
        prev_month, prev_count = results[i - 1]

        percent_change = ((current_count / prev_count) * 100) if prev_count > 0 else 0

        month_col = dbc.Col([
            html.Div(current_month.strftime('%b').upper()[:3], className="month-name text-center"),
            html.P("QTD. CLIENTES", className="info-text text-center", style={"fontSize": "11px"}),
            html.H6(f"{current_count}", className="info-number text-center", style={"fontSize": "17px"}),
            html.Div([
                html.Span("POSITIVAÇÃO", style={"display": "block", "fontSize": "10px", "color": "#A3AED0"}),
                html.Span(f"{percent_change:.2f}%", style={"color": "#3FB9C6", "display": "block", "fontSize": "10px"}), 
            ], className="info-percent text-center")
        ], width=2)

        children.append(month_col)

    month_row = dbc.Row(children, className="mb-4")

    return [month_row]

# #################### Card tabela Cliente Sintético
def preparar_dados_cliente_sintetico(vendedor_selecionado, df, ano_selecionado, visualizacao='total'):

    df = df.drop_duplicates(subset='Id_Pedido', keep='first')

    if vendedor_selecionado != "":
        df = df[df['Vendedor'] == vendedor_selecionado]

    # Garantir que 'Cliente_ID_Nome' esteja criado corretamente
    df['Cliente_ID_Nome'] = df['Matriz_Cliente'].astype(str) + ' - ' + df['Cliente']

    # Filtrar por ano
    df_filtrado = df[df['Data_Pedido'].dt.year == ano_selecionado]

    # Agrupar e somar os valores
    if visualizacao == 'total':
        df_agrupado = df_filtrado.groupby(['Cliente_ID_Nome', 'Cidade', df_filtrado['Data_Pedido'].dt.strftime('%m/%Y')])['TOTAL'].sum().unstack(fill_value=0)
    else:
        df_agrupado = df_filtrado.groupby(['Cliente_ID_Nome', 'Cidade', df_filtrado['Data_Pedido'].dt.strftime('%m/%Y')])['m2_pedido'].sum().unstack(fill_value=0)

    # Resetar índice para tornar as colunas 'Cliente_ID_Nome' e 'Cidade' parte do DataFrame
    df_agrupado.reset_index(inplace=True)

    # Inserir todos os meses como colunas, mesmo se não houver dados
    todos_os_meses = pd.date_range(start=f'{ano_selecionado}-01-01', end=f'{ano_selecionado}-12-31', freq='MS').strftime('%m/%Y').tolist()
    for mes in todos_os_meses:
        if mes not in df_agrupado.columns:
            df_agrupado[mes] = 0

    # Reordenar as colunas se necessário
    colunas = ['Cliente_ID_Nome', 'Cidade'] + todos_os_meses
    df_agrupado = df_agrupado[colunas]

    # Renomear colunas para o DataTable
    df_agrupado.columns = ['Cliente', 'Cidade'] + todos_os_meses

    return df_agrupado

# Callback para atualizar a tabela "Cliente Sintético"
@app.callback(
    Output("tabela-cliente-sintetico", "children"),
    [
        Input("interval-update", "n_intervals"),
        Input("filtro_ano", "value"),
        Input('id-busca-input', 'value'),
        Input('filtro_visualizacao', 'value'),
        Input('vendedor-dropdown', 'value')
    ] 
)
def update_tabela_cliente_sintetico(n_intervals, ano_selecionado, id_busca, visualizacao, vendedor_selecionado):
    df_cliente_sintetico = preparar_dados_cliente_sintetico(vendedor_selecionado, df, ano_selecionado, visualizacao)

    if id_busca:
        id_busca_str = str(id_busca)
        mask = df_cliente_sintetico['Cliente'].str.contains(id_busca_str, case=False, na=False)

        if mask.any():
            df_cliente_sintetico = pd.concat([df_cliente_sintetico[mask], df_cliente_sintetico[~mask]])
        else:
            return dbc.Alert(f'ID {id_busca} não encontrado.', color='danger')

    # Format values based on visualization type
    if visualizacao == 'metragem':
        for mes in df_cliente_sintetico.columns[2:]:
            df_cliente_sintetico[mes] = df_cliente_sintetico[mes].apply(lambda x: f"{x:.2f} m²" if x != 0 else x)
    else:  # Faturamento
        for mes in df_cliente_sintetico.columns[2:]:
            df_cliente_sintetico[mes] = df_cliente_sintetico[mes].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if x != 0 else x)

    columns = [
        {'name': 'Cliente', 'id': 'Cliente'}, 
        {'name': 'Cidade', 'id': 'Cidade'}
    ] + [
        {'name': mes, 'id': mes} 
        for mes in df_cliente_sintetico.columns[2:]
    ]

    tabela = dash_table.DataTable(
        data=df_cliente_sintetico.to_dict('records'),
        columns=columns,
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'center',
            'minWidth': '120px', 'width': '160px', 'maxWidth': '180px',
            'overflow': 'hidden',
            'textOverflow': 'ellipsis',
        },
        style_cell_conditional=[
            {'if': {'column_id': 'Cliente'}, 'textAlign': 'left', 'minWidth': '160px', 'width': '280px', 'maxWidth': '300px'},
            {'if': {'column_id': 'Cidade'}, 'textAlign': 'center', 'minWidth': '80px', 'width': '110px', 'maxWidth': '120px'},
        ],
        style_header={
            'fontWeight': 'bold',
            'textAlign': 'center'
        },
        page_size=20,
    )

    return tabela

def create_cliente_sintetico_card():
    return dbc.Card(
        [
            dbc.CardHeader(html.H3("CLIENTE SINTÉTICO", className="card-header-custom")),
            dbc.CardBody(
                [
                 dbc.Row(
                        [
                            dbc.Col(
                                dbc.Input(
                                    id='id-busca-input',
                                    type='number',
                                    placeholder='Buscar por ID do Cliente',
                                    style={'width': '150px','margin-bottom': '10px'}
                                ),
                                width="auto",
                            ),
                            dbc.Col(
                                [
                                    dcc.Dropdown(
                                        id='filtro_visualizacao',
                                        options=[
                                            {'label': 'Faturamento', 'value': 'total'},
                                            {'label': 'Metragem', 'value': 'metragem'}
                                        ],
                                        value='total',
                                        clearable=False,
                                        style={'width': '150px', 'margin-bottom': '10px'}
                                    ),
                                ],
                                width="auto",
                            ),
                            dbc.Col(
                                [
                                    dcc.Dropdown(
                                        id='filtro_ano',
                                        options=[{'label': ano, 'value': ano} for ano in range(datetime.now().year-1, datetime.now().year + 1)],
                                        value=datetime.now().year,
                                        clearable=False,
                                        style={'width': '120px', 'margin-bottom': '10px'}
                                    ),
                                ],
                                width="auto",
                            ),
                            dbc.Col(
                                [
                                    html.Button("Exportar Excel", id="btn_exportar", n_clicks=0, className="btn btn-custom"),
                                ],
                                width="auto",
                                style={"justify-content": "flex-end"},
                            ),
                        ],
                        justify="start",  # Isso vai alinhar os itens à esquerda
                    ),
                    html.Div(id="tabela-cliente-sintetico"),
                ]
            ),
        ],
        style={"width": "100%"},
    )
  
@app.callback(
    Output('download-excel', 'data'),
    [
        Input('btn_exportar', 'n_clicks'),
    ],
    [
        State('filtro_ano', 'value'),
        State('id-busca-input', 'value'),
        State('filtro_visualizacao', 'value'),
        State('vendedor-dropdown', 'value')
    ],
    prevent_initial_call=True
)
def exportar_para_excel(n_clicks, ano_selecionado, id_busca, visualizacao, vendedor_selecionado):
    if n_clicks > 0:
        df_cliente_sintetico = preparar_dados_cliente_sintetico(vendedor_selecionado, df, ano_selecionado, visualizacao)

        # Aplicar filtro de busca por ID se houver algum
        if id_busca:
            id_busca_str = str(id_busca)
            mask = df_cliente_sintetico['Cliente'].str.contains(id_busca_str, case=False, na=False)
            if mask.any():
                # Criar DataFrame concatenado com resultados filtrados primeiro e não filtrados em seguida
                df_cliente_sintetico = pd.concat([df_cliente_sintetico[mask], df_cliente_sintetico[~mask]])
            else:
                return dcc.send_bytes(b'', filename='nenhum_resultado.xlsx')  # Se não houver resultados, enviar um arquivo vazio ou com aviso.

        # Preparar o DataFrame para exportação conforme o tipo de visualização
        if visualizacao == 'metragem':
            for mes in df_cliente_sintetico.columns[2:]:
                df_cliente_sintetico[mes] = df_cliente_sintetico[mes].apply(lambda x: f"{x:.2f} m²" if x != 0 else x)
        else:  # Faturamento
            for mes in df_cliente_sintetico.columns[2:]:
                df_cliente_sintetico[mes] = df_cliente_sintetico[mes].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if x != 0 else x)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_cliente_sintetico.to_excel(writer, index=False)
        output.seek(0)

        # Enviar o buffer para download
        return dcc.send_bytes(output.getvalue(), filename=f"clientes_sinteticos_{ano_selecionado}.xlsx")

    return None

# Pegar o os dados para os cards
valor_realizado = calc_realizado(df)
vendedor_names = get_vendedor_names(df)  
realizado_ate_ontem = calc_realizado_ate_ontem(df)
valor_projetado = calc_projecao_geral(realizado_ate_ontem)
cliente_sintetico_card = create_cliente_sintetico_card()
categorias_agregadas = ['ACESSÓRIOS', 'ALUMÍNIO', 'FERRAGEM', 'KIT PARA BOX PADRÃO', 'SILICONE']
categoria_vidro = ['VIDRO']

# Calcular a projeção pela meta geral
projecao_pela_meta_geral = (valor_projetado / meta_geral_valor) * 100

# Determinar o percentual de comissão com base na projeção pela meta geral
if projecao_pela_meta_geral >= 100:
    percentual_comissao = 1.3
elif projecao_pela_meta_geral >= 95:
    percentual_comissao = 1.2
elif projecao_pela_meta_geral >= 90:
    percentual_comissao = 1.1
else:
    percentual_comissao = 1.0

# Formatar o percentual de comissão com uma casa decimal
percentual_comissao_str = f"{percentual_comissao:.1f}%"
# Definição do tooltip
tooltip_text = f"""
100% da Meta Geral: 1.3%
95% da Meta Geral: 1.2%
90% da Meta Geral: 1.1%
Valor atual: {projecao_pela_meta_geral:.2f}%
"""

# Preparando os dados
df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')
hoje = datetime.now()
meses = []

# Calculando os limites dos últimos três meses
for i in range(2, 5):  
    inicio_mes = (hoje - pd.offsets.MonthBegin(n=i)).to_pydatetime()
    fim_mes = (inicio_mes + pd.offsets.MonthEnd(n=0)).to_pydatetime()
    meses.append((inicio_mes, fim_mes))

meses.reverse()

# Supondo que 'TODOS OS VENDEDORES' seja o valor padrão para incluir todos os vendedores
vendedor_padrao = "TODOS OS VENDEDORES"

# Calculando as somas para cada mês e categoria com o valor padrão para vendedor
somas_agregadas = [calcular_somas(df, categorias_agregadas, inicio, fim, vendedor_padrao) for inicio, fim in meses]
somas_vidro = [calcular_somas(df, categoria_vidro, inicio, fim, vendedor_padrao) for inicio, fim in meses]

nomes_meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Criando o gráfico de barras
fig_pilha = go.Figure(data=[
    go.Bar(name='Vidro', x=nomes_meses, y=somas_vidro, marker_color='#3FB9C6', width=0.4, textposition='auto'),
    go.Bar(name='Agregadas', x=nomes_meses, y=somas_agregadas, marker_color='#8F9BBA', width=0.4, textposition='auto')
])

# Atualizando o layout do gráfico para empilhar
fig_pilha.update_layout(barmode='stack', title='VENDAS POR CATEGORIA ÚLTIMOS 3 MESES')

# Callback para atualizar o gráfico com base no vendedor selecionado
@app.callback(
    Output('VENDAS POR CATEGORIA ÚTIMOS 3 MESES', 'figure'),
    [Input('vendedor-dropdown', 'value')]
)
def update_graph(vendedor_selecionado):
    # Calculando as somas para o vendedor selecionado
    somas_agregadas = [calcular_somas(df, categorias_agregadas, inicio, fim, vendedor_selecionado) for inicio, fim in meses]
    somas_vidro = [calcular_somas(df, categoria_vidro, inicio, fim, vendedor_selecionado) for inicio, fim in meses]
    # Criando o gráfico atualizado
    fig_pilha = go.Figure(data=[
    go.Bar(
        name='Vidro',
        x=nomes_meses,
        y=somas_vidro,
        marker_color='#3FB9C6',
        width=0.4,
        textposition='outside'  # Posicionamento do texto fora da coluna
    ),
    go.Bar(
        name='Agregadas',
        x=nomes_meses,
        y=somas_agregadas,
        marker_color='#8F9BBA',
        width=0.4,
        textposition='outside'  # Posicionamento do texto fora da coluna
    )
    ])

        # Alterar a disposição do gráfico para empilhar
    fig_pilha.update_layout(
        barmode='stack',
        title='VENDAS POR CATEGORIA ÚLTIMOS 3 MESES',
        xaxis=dict(title='Mês'),
        yaxis=dict(title='Vendas'),
        margin=dict(l=20, r=20, t=40, b=20)  # Ajusta as margens se necessário para evitar cortes de texto
    )

    return fig_pilha

df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')
##################### Card tabela faturamento vidro 3 meses
def calcular_somas_grupos(df, subcategoria):
    # Obtém o primeiro e o último dia do mês atual
    hoje = pd.to_datetime('today').normalize()
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = primeiro_dia_mes + pd.offsets.MonthEnd(1)
    df_filtrado = df[(df['Data_Pedido'] >= primeiro_dia_mes) & 
                     (df['Data_Pedido'] <= ultimo_dia_mes) & 
                     (df['Tipo_Produto'] == subcategoria)]

    # Aplica o desconto antes de somar
    df_filtrado['total_produto_com_desconto'] = df_filtrado.apply(apply_discount, axis=1)

    # Retorna a soma dos totais com desconto
    return df_filtrado['total_produto_com_desconto'].sum()

def calcular_somas_grupos_frete(df):
    df['PERIODO'] = pd.to_datetime(df['PERIODO'])
    # Obtém o primeiro e o último dia do mês atual
    hoje = pd.to_datetime('today').normalize()
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = primeiro_dia_mes + pd.offsets.MonthEnd(1)
    df_filtrado = df[(df['PERIODO'] >= primeiro_dia_mes) & 
                     (df['PERIODO'] <= ultimo_dia_mes)]
    return df_filtrado['Frete'].sum()

def calcular_somas_grupos_benef(df):
    df['PERIODO'] = pd.to_datetime(df['PERIODO'])
    # Obtém o primeiro e o último dia do mês atual
    hoje = pd.to_datetime('today').normalize()
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = primeiro_dia_mes + pd.offsets.MonthEnd(1)
    df_filtrado = df[(df['PERIODO'] >= primeiro_dia_mes) & 
                     (df['PERIODO'] <= ultimo_dia_mes)]
    return df_filtrado['FATURAMENTO'].sum()

@app.callback(
    Output('faturamento_vidro_card_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_faturamento_vidro_card(vendedor_selecionado):
    if not vendedor_selecionado:
        # Se por algum motivo o vendedor_selecionado for None, use o valor padrão
        vendedor_selecionado = "TODOS OS VENDEDORES"
    
    return create_faturamento_vidro_card(df, vendedor_selecionado)

icone_svg = """![icone](assets/img/topmes.svg)"""

def format_currency(value):
    """Formata valores como moeda brasileira."""
    return f"R$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

# Função para criar o card da tabela
def create_faturamento_vidro_card(df, vendedor_selecionado):
    return dbc.Card([
        dbc.CardHeader(
                html.H3("FATURAMENTO DOS ÚLTIMOS 3 MESES VIDRO"),
                className="card-header-custom",
            ),
        dbc.CardBody(
            html.Div(
                create_faturamento_vidro_table(df, vendedor_selecionado)
            )
        ),
    ])

def create_faturamento_vidro_table(df, vendedor_selecionado):
    subcategorias = {
        'TEMPERADO ENGENHARIA': ['ENGENHARIA TEMPERADO', 'BOX ENGENHARIA'],
        'TEMPERADO PRONTA ENTREGA': ['BOX PADRÃO', 'JANELA PADRÃO', 'PORTA PIVOTANTE'],
        'COMUM CORTADO': ['CORTADO ESPELHO', 'CORTADO FLOAT', 'CORTADO LAMINADO', 'CORTADO FANTASIA', 'CORTADO REFLETIVO BRONZE', 'CORTADO SERIGRAFADO'],
        'COMUM CHAPARIA': ['CHAPARIA ESPELHO', 'CHAPARIA FANTASIA', 'CHAPARIA FLOAT', 'CHAPARIA LAMINADO', 'CHAPARIA REFLETIVO BRONZE', 'CHAPARIA SERIGRAFADO'],
    }

    # Filtrar por vendedor, se não for "TODOS OS VENDEDORES"
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]

    totais = {'Subcategoria': 'Totais'}
    total_values = {}

    for start_date in start_dates:
        total_values[start_date.strftime('%b/%Y')] = 0
    
    data = []
    for grupo, tipos in subcategorias.items():
        grupo_row = {'Subcategoria': grupo}
        grupo_somas = {start_date.strftime('%b/%Y'): 0 for start_date in start_dates}

        for tipo in tipos:
            tipo_row = {'Subcategoria': f"· {tipo}"}
            faturamento_mes = {}
            for start_date in start_dates:
                fim_mes = start_date + pd.offsets.MonthEnd(1)
                df_filtrado = df[(df['Data_Pedido'] >= start_date) & (df['Data_Pedido'] <= fim_mes) & (df['Tipo_Produto'] == tipo)]
                df_filtrado['Faturamento_com_Desconto'] = df_filtrado.apply(apply_discount, axis=1)
                faturamento = df_filtrado['Faturamento_com_Desconto'].sum()
                tipo_row[start_date.strftime('%b/%Y')] = format_currency(faturamento)
                faturamento_mes[start_date.strftime('%b/%Y')] = faturamento
                grupo_somas[start_date.strftime('%b/%Y')] += faturamento
                total_values[start_date.strftime('%b/%Y')] += faturamento

            # Determina o TOP MÊS para a subcategoria
            top_mes_faturamento = max(faturamento_mes.values())
            top_mes_subcategoria = [mes for mes, faturamento in faturamento_mes.items() if faturamento == top_mes_faturamento][0]
            tipo_row['TOP MÊS'] = top_mes_subcategoria
            
            data.append(tipo_row)
        
        # Adiciona faturamento total e TOP MÊS ao grupo
        for mes, soma in grupo_somas.items():
            grupo_row[mes] = f"R$ {soma:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        top_mes_grupo_faturamento = max(grupo_somas.values())
        grupo_row['TOP MÊS'] = [mes for mes, faturamento in grupo_somas.items() if faturamento == top_mes_grupo_faturamento][0]
        data.insert(len(data) - len(tipos), grupo_row)

    # Adicionar a categoria "Outros Vidros" ao final
    row_outros_vidros = {'Subcategoria': 'OUTROS VIDROS'}
    soma_outros_vidros = {}
    for start_date in start_dates:
        fim_mes = start_date + pd.offsets.MonthEnd(1)
        df_filtrado = df[(df['Data_Pedido'] >= start_date) & (df['Data_Pedido'] <= fim_mes) & (~df['Tipo_Produto'].isin(sum(subcategorias.values(), []))) & (df['Grupo'] == 'VIDRO')]
        df_filtrado['Faturamento_com_Desconto_'] = df_filtrado.apply(apply_discount, axis=1)
        soma_grupo = df_filtrado['Faturamento_com_Desconto_'].sum()
        row_outros_vidros[start_date.strftime('%b/%Y')] = f"R$ {soma_grupo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        soma_outros_vidros[start_date.strftime('%b/%Y')] = soma_grupo
        total_values[start_date.strftime('%b/%Y')] += soma_grupo
    
    # Determinar o TOP MÊS para "OUTROS VIDROS"
    top_mes_outros_vidros = max(soma_outros_vidros, key=soma_outros_vidros.get)
    row_outros_vidros['TOP MÊS'] = top_mes_outros_vidros
    data.append(row_outros_vidros)

    top_mes_totais = max(total_values, key=total_values.get)
    totais['TOP MÊS'] = top_mes_totais
    for mes, total in total_values.items():
        totais[mes] = format_currency(total)
    data.insert(0, totais)
    
    columns = [
        {"name": "Subcategoria", "id": "Subcategoria", "type": "text"}
    ] + [
        {"name": mes.strftime('%b/%Y'), "id": mes.strftime('%b/%Y'), "type": "numeric", "format": Format(symbol=Symbol.yes, symbol_suffix='R$ ', scheme=Scheme.fixed)}
        for mes in start_dates
    ] + [
        {"name": "TOP MÊS", "id": "TOP MÊS", "type": "text","presentation": "markdown"}
    ]

    style_cell = {
    'textAlign': 'left',
    'padding': '0px',  
    'whiteSpace': 'normal',
    'margin': '-15%' 
    }

    style_header = {
        'fontWeight': 'bold',
        'textAlign': 'left',
    }

    return html.Div(children=[
    dash_table.DataTable(
        data=data,
        columns=columns,
        style_cell=style_cell,
        style_header=style_header,
        style_cell_conditional=[
                {
                    'if': {'column_id': 'Subcategoria'},
                    'width': '8%',
                    'whiteSpace': 'normal' 
                }
            ],
       style_data_conditional=[
        {'if': {'row_index': 0}, 'backgroundColor': '#F1F1F1', 'fontWeight': 'bold'},  
        {
            'if': {'filter_query': '{Subcategoria} = "TEMPERADO ENGENHARIA"'},
            'backgroundColor': '#b0c4de',  # Light steel blue
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "TEMPERADO PRONTA ENTREGA"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "COMUM CORTADO"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "COMUM CHAPARIA"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "OUTROS VIDROS"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        }])
])

def calcular_volume_por_categoria(df, categoria):
    # Obtém o primeiro e o último dia do mês atual
    hoje = pd.to_datetime('today').normalize()
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = primeiro_dia_mes + pd.offsets.MonthEnd(1)

    df_filtrado = df[(df['Data_Pedido'] >= primeiro_dia_mes) & 
                     (df['Data_Pedido'] <= ultimo_dia_mes) & 
                     (df['Tipo_Produto'] == categoria)]

    # Calcula e retorna a soma da coluna 'm2' para as linhas filtradas
    return df_filtrado['m2'].sum()

def calc_projecao_categoria(realizado):
    hoje = pd.to_datetime('today')
    ontem = hoje - pd.offsets.BDay(1)  # Calcula o dia útil anterior
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = hoje + pd.offsets.MonthEnd(1)
    
    dias_uteis_ate_ontem = pd.bdate_range(start=primeiro_dia_mes, end=ontem).size
    total_dias_uteis_mes = pd.bdate_range(start=primeiro_dia_mes, end=ultimo_dia_mes).size

    if dias_uteis_ate_ontem > 0:
        projecao = (realizado / dias_uteis_ate_ontem) * total_dias_uteis_mes
        return projecao
    return 0

def calcular_media_faturamento_ultimos_3_meses(df, subcategorias):
    def calcular_somas_grupos_(df, subcategoria, inicio_mes, fim_mes):
        df_filtrado = df[(df['Data_Pedido'] >= inicio_mes) & (df['Data_Pedido'] <= fim_mes) & (df['Tipo_Produto'].isin([subcategoria]))]
        df_filtrado['total_produto_com_desconto'] = df_filtrado.apply(apply_discount, axis=1)
        return df_filtrado['total_produto_com_desconto'].sum()
    
    # Calcula a média do faturamento dos últimos 3 meses
    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]
    faturamento_total = sum(
        calcular_somas_grupos_(df, tipo, start_date, start_date + pd.offsets.MonthEnd(1))
        for start_date in start_dates
        for tipo in subcategorias
    ) 
    return (faturamento_total / 3)

def calcular_media_faturamento_ultimos_3_meses_frete(df):
    df['PERIODO'] = pd.to_datetime(df['PERIODO'])
    def calcular_somas_grupos_(df, inicio_mes, fim_mes):
        df_filtrado = df[(df['PERIODO'] >= inicio_mes) & (df['PERIODO'] <= fim_mes)]
        return df_filtrado['Frete'].sum()
    
    # Calcula a média do faturamento dos últimos 3 meses
    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]
    faturamento_total = sum(
        calcular_somas_grupos_(df, start_date, start_date + pd.offsets.MonthEnd(1))
        for start_date in start_dates
    ) 
    return (faturamento_total / 3)

def calcular_media_faturamento_ultimos_3_meses_benef(df):
    df['PERIODO'] = pd.to_datetime(df['PERIODO'])
    def calcular_somas_grupos_(df, inicio_mes, fim_mes):
        df_filtrado = df[(df['PERIODO'] >= inicio_mes) & (df['PERIODO'] <= fim_mes)]
        return df_filtrado['FATURAMENTO'].sum()

    # Calcula a média do faturamento dos últimos 3 meses
    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]
    faturamento_total = sum(
        calcular_somas_grupos_(df, start_date, start_date + pd.offsets.MonthEnd(1))
        for start_date in start_dates
    ) 
    return (faturamento_total / 3)

@app.callback(
    Output('categoria_vidro_table_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_categoria_vidro_table(vendedor_selecionado):
    if not vendedor_selecionado:
        vendedor_selecionado = "TODOS OS VENDEDORES"

    return create_categoria_vidro_card(df, vendedor_selecionado)

def create_categoria_vidro_card(df, vendedor_selecionado):
    return dbc.Card(
        [
            dbc.CardHeader(
                html.H3("CATEGORIA VIDRO"),
                className="card-header-custom",
            ),
            dbc.CardBody(
                create_categoria_vidro_table(df, vendedor_selecionado)
                ),
        ])

def create_categoria_vidro_table(df, vendedor_selecionado):
    subcategorias = {
        'TEMPERADO ENGENHARIA': ['ENGENHARIA TEMPERADO', 'BOX ENGENHARIA'],
        'TEMPERADO PRONTA ENTREGA': ['BOX PADRÃO', 'JANELA PADRÃO', 'PORTA PIVOTANTE'],
        'COMUM CORTADO': ['CORTADO ESPELHO', 'CORTADO FLOAT', 'CORTADO LAMINADO', 'CORTADO FANTASIA', 'CORTADO REFLETIVO BRONZE', 'CORTADO SERIGRAFADO'],
        'COMUM CHAPARIA': ['CHAPARIA ESPELHO', 'CHAPARIA FANTASIA', 'CHAPARIA FLOAT', 'CHAPARIA LAMINADO', 'CHAPARIA REFLETIVO BRONZE', 'CHAPARIA SERIGRAFADO'],
    }

    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    totals = {'Subcategoria': 'Totais', 'Volume': 0, 'Realizado': 0, 'Projeção': 0, 'Meta': 0}
    data = []

    for grupo, tipos in subcategorias.items():
        grupo_data = {'Subcategoria': grupo, 'Volume': 0, 'Realizado': 0, 'Projeção': 0, 'Meta': 0}
        grupo_total_realizado = 0
        grupo_total_meta = 0

        for tipo in tipos:
            volume = calcular_volume_por_categoria(df, tipo)
            realizado = calcular_somas_grupos(df, tipo)
            projecao = calc_projecao_categoria(realizado)
            meta = calcular_media_faturamento_ultimos_3_meses(df, [tipo])

            # Atualiza somatórios do grupo
            grupo_data['Volume'] += volume
            grupo_data['Realizado'] += realizado
            grupo_data['Projeção'] += projecao
            grupo_data['Meta'] += meta
            grupo_total_realizado += realizado
            grupo_total_meta += meta

            # Insere dados da subcategoria
            data.append({
                'Subcategoria': f"· {tipo}",
                'Volume': f"{volume:.2f} m²",
                'Realizado': "R$ {:,.2f}".format(realizado).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção': "R$ {:,.2f}".format(projecao).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Meta': "R$ {:,.2f}".format(meta).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção vs Meta': f"{(projecao / meta * 100) if meta else 0:.2f}%",
            })
        # Acumula totais agregados
        totals['Volume'] += grupo_data['Volume']
        totals['Realizado'] += grupo_data['Realizado']
        totals['Projeção'] += grupo_data['Projeção']
        totals['Meta'] += grupo_data['Meta']

        # Calcula e adiciona projeção vs meta para o grupo
        grupo_data['Projeção vs Meta'] = f"{(grupo_data['Projeção'] / grupo_data['Meta'] * 100) if grupo_data['Meta'] else 0:.2f}%"

        # Formata valores para o grupo
        grupo_data['Volume'] = f"{grupo_data['Volume']:.2f} m²"
        grupo_data['Realizado'] = "R$ {:,.2f}".format(grupo_data['Realizado']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Projeção'] = "R$ {:,.2f}".format(grupo_data['Projeção']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Meta'] = "R$ {:,.2f}".format(grupo_data['Meta']).replace(',', 'X').replace('.', ',').replace('X', '.')

        data.insert(len(data) - len(tipos), grupo_data)
    df_outros_vidros = df[~df['Tipo_Produto'].isin(sum(subcategorias.values(), [])) & (df['Grupo'] == 'VIDRO')]
 
    volume_outros = calcular_volume_por_categoria(df_outros_vidros, "OUTROS VIDROS")
    realizado_outros = calcular_somas_grupos(df_outros_vidros, "OUTROS VIDROS")
    projecao_outros = calc_projecao_categoria(realizado_outros)
    meta_outros = calcular_media_faturamento_ultimos_3_meses(df_outros_vidros, ["OUTROS VIDROS"])
    projecao_vs_meta_outros = f"{(projecao_outros / meta_outros) * 100:.2f}%" if meta_outros > 0 else "0.00%"

    # Adicionando "OUTROS VIDROS" aos dados
    data.append({
        'Subcategoria': 'OUTROS VIDROS',
        'Volume': f"{volume_outros:,.2f} m²",
        'Realizado': "R$ {:,.2f}".format(realizado_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção': "R$ {:,.2f}".format(projecao_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Meta': "R$ {:,.2f}".format(meta_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção vs Meta': projecao_vs_meta_outros
    })

    for key in ['Realizado', 'Projeção', 'Meta']:
        totals[key] = "R$ {:,.2f}".format(totals[key]).replace(',', 'X').replace('.', ',').replace('X', '.')
    for key in ['Volume']:
        totals[key] = "{:,.2f} m²".format(totals[key]).replace(',', 'X').replace('.', ',').replace('X', '.')

    data.insert(0, totals)

    style_cell = {
    'textAlign': 'left',
    'padding': '0px',  
    'whiteSpace': 'normal',
    }

    style_header = {
        'fontWeight': 'bold',
        'textAlign': 'left',
    }

    columns = [
        {'name': 'Subcategoria', 'id': 'Subcategoria'}
    ] + [
        {'name': 'Volume', 'id': 'Volume'},
    ] + [
        {'name': 'Realizado', 'id': 'Realizado'},
    ] + [
        {'name': 'Projeção', 'id': 'Projeção'},
    ] + [
        {'name': 'Meta', 'id': 'Meta', 'type': 'numeric', 'format': Format(symbol=Symbol.yes, symbol_suffix='R$ ', scheme=Scheme.fixed)},
    ] + [
        {'name': 'Projeção vs Meta', 'id': 'Projeção vs Meta'}
    ]

    return html.Div(children=[
        dash_table.DataTable(
        columns=columns,
        data=data,
        style_cell=style_cell,
        style_header=style_header,
        style_cell_conditional=[
                {
                    'if': {'column_id': 'Subcategoria'},
                    'width': '8%',
                    'whiteSpace': 'normal' 
                }
            ],
       style_data_conditional=[
        {'if': {'row_index': 0}, 'backgroundColor': '#F1F1F1', 'fontWeight': 'bold'},
        {
            'if': {'filter_query': '{Subcategoria} = "TEMPERADO ENGENHARIA"'},
            'backgroundColor': '#b0c4de',  
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "TEMPERADO PRONTA ENTREGA"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "COMUM CORTADO"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "COMUM CHAPARIA"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "OUTROS VIDROS"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        }])
])


################### TABELA FATURAMENTO DOS ÚLTIMOS 3 MESES DE AGREGADOS
@app.callback(
    Output('faturamento_agregados_3m_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_faturamento_agregados_table(vendedor_selecionado):
    if not vendedor_selecionado:
        vendedor_selecionado = "TODOS OS VENDEDORES"

    return create_categoria_vidro_agregados_card(df, vendedor_selecionado)

def create_categoria_vidro_agregados_card(df, vendedor_selecionado):
    return dbc.Card([
            dbc.CardHeader(
                html.H3("FATURAMENTO DOS ÚLTIMOS 3 MESES AGREGADOS"),
                className="card-header-custom",
            ),
            dbc.CardBody(
                create_faturamento_agregados_table(df, vendedor_selecionado),
              style={'margin-bottom': '0px', 'padding-bottom': '0px'}),
            ])

def apply_discount_benef(row):
    if row['Tipo_Desconto'] == 'Porcentagem':
        return max(0, row['valor_beneficiamento'] - (row['valor_beneficiamento'] * row['Desconto'] / 100))
    elif row['Tipo_Desconto'] == 'Reais':
        valor_frete = row.get('Valor_Frete', 0)
        return ((row['valor_beneficiamento'] * row['Desconto']) / ((row['TOTAL'] - valor_frete) + row['Desconto']) - row['valor_beneficiamento']) * (-1)
    return row['valor_beneficiamento'] 

def create_faturamento_agregados_table(df, vendedor_selecionado):
    df_frete = fetch_data_frete()
    df_benef = fetch_data_benef()
    df_frete['Vendedor'] = df_frete['Vendedor'].str.split().str.get(0)
    df_benef['Vendedor'] = df_benef['Vendedor'].str.split().str.get(0)
    df_madeira = df_benef[df_benef['NOME_BENEF'] == 'Caixa de Madeira']
    df_benef = df_benef[df_benef['NOME_BENEF'] != 'Caixa de Madeira']

    subcategorias = {
        'KIT\'S PARA BOX': ['KIT BOX COMPLETO AL', 'KIT BOX COMPLETO IDEIA GLASS', 'KIT BOX COMPLETO IMPORTADO', 'KIT BOX COMPLETO PORTAL', 'KIT BOX COMPLETO PORTAL - AVARIA'],
        'KIT\'S PARA JANELA': ['KIT JANELA COMPLETA WD', 'KIT JANELA COMPLETO PORTAL'],
        'PERFIS PARA VIDRO TEMPERADO': ['PERFIS ENGENHARIA AL', 'PERFIS ENGENHARIA PERFILEVE', 'PERFIS ENGENHARIA PERFILEVE 3MTS'],
        'FERRAGENS': ['KIT FERRAGENS LGL', 'FERRAGENS LGL', 'MOLAS', 'ROLDANAS', 'PUXADORES'],
        'OUTROS': ['SILICONE', 'FIXA ESPELHO', 'SUPORTES', 'BORRACHAS', 'ESCOVINHAS'],
        'SERVIÇOS': ['MÃO DE OBRA'],
    }
    
    # Filtrando dados conforme o vendedor selecionado
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]
        df_frete = df_frete[df_frete['Vendedor'] == vendedor_selecionado]
        df_benef = df_benef[df_benef['Vendedor'] == vendedor_selecionado]
        df_madeira = df_madeira[df_madeira['Vendedor'] == vendedor_selecionado]

    # Convertendo datas
    df_frete['PERIODO'] = pd.to_datetime(df_frete['PERIODO'])
    df_benef['PERIODO'] = pd.to_datetime(df_benef['PERIODO'])
    df_madeira['PERIODO'] = pd.to_datetime(df_madeira['PERIODO'])

    # Definindo as datas de início de cada mês para os últimos 3 meses
    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]

    totais = {'Subcategoria': 'Totais'}
    total_values = {}
    for start_date in start_dates:
        total_values[start_date.strftime('%b/%Y')] = 0
    data = []
    for grupo, tipos in subcategorias.items():
        grupo_row = {'Subcategoria': grupo}
        grupo_somas = {start_date.strftime('%b/%Y'): 0 for start_date in start_dates}
        data.append(grupo_row)

        for tipo in tipos:
            tipo_row = {'Subcategoria': f"· {tipo}"}
            faturamento_mes = {}
            
            for start_date in start_dates:
                fim_mes = start_date + pd.offsets.MonthEnd(1)
                df_tipo_filtrado = df[
                    (df['Data_Pedido'] >= start_date) &
                    (df['Data_Pedido'] <= fim_mes) &
                    (df['Tipo_Produto'] == tipo)
                ]
                df_tipo_filtrado['Faturamento_com_Desconto'] = df_tipo_filtrado.apply(apply_discount, axis=1)
                faturamento = df_tipo_filtrado['Faturamento_com_Desconto'].sum()
                tipo_row[start_date.strftime('%b/%Y')] = format_currency(faturamento)
                faturamento_mes[start_date.strftime('%b/%Y')] = faturamento
                grupo_somas[start_date.strftime('%b/%Y')] += faturamento
                total_values[start_date.strftime('%b/%Y')] += faturamento

            top_mes_faturamento = max(faturamento_mes.values())
            top_mes_subcategoria = [mes for mes, faturamento in faturamento_mes.items() if faturamento == top_mes_faturamento][0]
            tipo_row['TOP MÊS'] = top_mes_subcategoria
            data.append(tipo_row)

        for mes, soma in grupo_somas.items():
            grupo_row[mes] = f"R$ {soma:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        top_mes_grupo_faturamento = max(grupo_somas.values())
        grupo_row['TOP MÊS'] = [mes for mes, faturamento in grupo_somas.items() if faturamento == top_mes_grupo_faturamento][0]

    # Processamento para Frete
    frete_somas = {}
    for start_date in start_dates:
        fim_mes = start_date + pd.offsets.MonthEnd(1)
        df_frete_filtrado = df_frete[(df_frete['PERIODO'] >= start_date) & (df_frete['PERIODO'] <= fim_mes)]
        frete_somas[start_date.strftime('%b/%Y')] = df_frete_filtrado['Frete'].sum()
        
    frete_row = {'Subcategoria': '· FRETE'}
    for mes, soma in frete_somas.items():
        frete_row[mes] = f"R$ {soma:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    top_mes_frete = max(frete_somas, key=frete_somas.get)
    frete_row['TOP MÊS'] = top_mes_frete
    data.append(frete_row)

    # Processamento para Beneficiamento
    beneficiamento_somas = {}
    for start_date in start_dates:
        fim_mes = start_date + pd.offsets.MonthEnd(1)
        df_filtrado = df_benef[(df_benef['PERIODO'] >= start_date) & (df_benef['PERIODO'] <= fim_mes)]
        beneficiamento_somas[start_date.strftime('%b/%Y')] = df_filtrado['FATURAMENTO'].sum()

    beneficiamento_row = {'Subcategoria': '· BENEFICIAMENTO'}
    for mes, soma in beneficiamento_somas.items():
        beneficiamento_row[mes] = f"R$ {soma:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    top_mes_beneficiamento = max(beneficiamento_somas, key=beneficiamento_somas.get)
    beneficiamento_row['TOP MÊS'] = top_mes_beneficiamento
    data.append(beneficiamento_row)

    # Processamento para Caixa de Madeira
    madeira_somas = {}
    for start_date in start_dates:
        fim_mes = start_date + pd.offsets.MonthEnd(1)
        df_filtrado_madeira = df_madeira[(df_madeira['PERIODO'] >= start_date) & (df_madeira['PERIODO'] <= fim_mes)]
        madeira_somas[start_date.strftime('%b/%Y')] = df_filtrado_madeira['FATURAMENTO'].sum()

    madeira_row = {'Subcategoria': '· CAIXA DE MADEIRA'}
    for mes, soma in madeira_somas.items():
        madeira_row[mes] = f"R$ {soma:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    top_madeira = max(madeira_somas, key=madeira_somas.get)
    madeira_row['TOP MÊS'] = top_madeira
    data.append(madeira_row)

    top_mes_totais = max(total_values, key=total_values.get)
    totais['TOP MÊS'] = top_mes_totais
    for mes, total in total_values.items():
        totais[mes] = format_currency(total)
    data.insert(0, totais)
    
    columns = [
        {"name": "Subcategoria", "id": "Subcategoria", "type": "text"}
    ] + [
        {"name": mes.strftime('%b/%Y'), "id": mes.strftime('%b/%Y'), "type": "numeric", "format": Format(symbol=Symbol.yes, symbol_suffix='R$ ', scheme=Scheme.fixed)}
        for mes in start_dates
    ] + [
        {"name": "TOP MÊS", "id": "TOP MÊS", "type": "text", "presentation": "markdown"}
    ]

    style_cell = {
    'textAlign': 'left',
    'padding': '0px',
    'whiteSpace': 'normal',
    'margin': '-15%' 
    }

    style_header = {
        'fontWeight': 'bold',
        'textAlign': 'left',
    }

    return html.Div([
        dash_table.DataTable(
            data=data,
            columns=columns,
            style_table={'overflowX': 'auto'},
            style_cell=style_cell,
            style_header=style_header,
            style_cell_conditional=[
                {
                    'if': {'column_id': 'Subcategoria'},
                    'width': '8%',
                    'whiteSpace': 'normal' 
                }
            ],
             style_data_conditional=[
        {'if': {'row_index': 0}, 'backgroundColor': '#F1F1F1', 'fontWeight': 'bold'},
        {
            'if': {'filter_query': '{Subcategoria} = "KIT\'S PARA BOX"'},
            'backgroundColor': '#b0c4de',  
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "KIT\'S PARA JANELA"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "PERFIS PARA VIDRO TEMPERADO"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "FERRAGENS"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "OUTROS"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        },
        {
            'if': {'filter_query': '{Subcategoria} = "SERVIÇOS"'},
            'backgroundColor': '#b0c4de',
            'color': 'black'
        }])
    ])

################### TABELA CATEGORIA AGREGADOS
@app.callback(
    Output('categoria_agregados_table_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_categoria_agregados_table(vendedor_selecionado):
    if not vendedor_selecionado:
        vendedor_selecionado = "TODOS OS VENDEDORES"
  
    return create_categorias_agregados_card(df, vendedor_selecionado)

def create_categorias_agregados_card(df, vendedor_selecionado):
    return dbc.Card(
        [
            dbc.CardHeader(
                html.H3("CATEGORIA AGREGADAS"),
                className="card-header-custom",
            ),
            dbc.CardBody(
                create_categoria_agregadas_table(df, vendedor_selecionado),
               style={'margin-bottom': '0px', 'padding-bottom': '0px'}),
        ])

def calcular_faturamento_frete(df_frete):
    # Filtra os dados dos últimos 3 meses
    df_frete['PERIODO'] = pd.to_datetime(df_frete['PERIODO'])
    ultimo_mes = df_frete['PERIODO'].max()
    tres_meses_atras = ultimo_mes - pd.DateOffset(months=3)
    df_frete_3_meses = df_frete[df_frete['PERIODO'] > tres_meses_atras]

    # Calcula a soma do faturamento dos últimos 3 meses
    faturamento_frete = df_frete_3_meses['Frete'].sum()
    return faturamento_frete

def create_categoria_agregadas_table(df, vendedor_selecionado):
    df_frete = fetch_data_frete()
    df_benef = fetch_data_benef()
    df_frete['Vendedor'] = df_frete['Vendedor'].str.split().str.get(0)
    df_benef['Vendedor'] = df_benef['Vendedor'].str.split().str.get(0)
    df_madeira = df_benef[df_benef['NOME_BENEF'] == 'Caixa de Madeira']
    df_benef = df_benef[df_benef['NOME_BENEF'] != 'Caixa de Madeira']

    subcategorias = {
    'KIT\'S PARA BOX': ['KIT BOX COMPLETO AL', 'KIT BOX COMPLETO IDEIA GLASS', 'KIT BOX COMPLETO IMPORTADO', 'KIT BOX COMPLETO PORTAL', 'KIT BOX COMPLETO PORTAL - AVARIA'],
    'KIT\'S PARA JANELA': ['KIT JANELA COMPLETA WD', 'KIT JANELA COMPLETO PORTAL'],
    'PERFIS PARA VIDRO TEMPERADO': ['PERFIS ENGENHARIA AL', 'PERFIS ENGENHARIA PERFILEVE', 'PERFIS ENGENHARIA PERFILEVE 3MTS'],
    'FERRAGENS': ['KIT FERRAGENS LGL', 'FERRAGENS LGL', 'MOLAS', 'ROLDANAS', 'PUXADORES'],
    'OUTROS': ['SILICONE', 'FIXA ESPELHO', 'SUPORTES', 'BORRACHAS', 'ESCOVINHAS'],
    'SERVIÇOS': ['MÃO DE OBRA'],
    }

    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]
        df_frete = df_frete[df_frete['Vendedor'] == vendedor_selecionado]
        df_benef = df_benef[df_benef['Vendedor'] == vendedor_selecionado]
        df_madeira = df_madeira[df_madeira['Vendedor'] == vendedor_selecionado]
    
    # Preparação da lista de dados e somatórios de grupo
    total_realizado, total_projecao, total_meta = 0, 0, 0
    data = []
    for grupo, tipos in subcategorias.items():
        grupo_data = {'Subcategoria': grupo, 'Volume': 0, 'Realizado': 0, 'Projeção': 0, 'Meta': 0}
        for tipo in tipos:
            realizado = calcular_somas_grupos(df, tipo)
            projecao = calc_projecao_categoria(realizado)
            meta = calcular_media_faturamento_ultimos_3_meses(df, [tipo])

            grupo_data['Realizado'] += realizado
            grupo_data['Projeção'] += projecao
            grupo_data['Meta'] += meta

            data.append({
                'Subcategoria': f"· {tipo}",
                'Realizado': f"R$ {realizado:,.2f}".format(grupo_data['Realizado']).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção': f"R$ {projecao:,.2f}".format(grupo_data['Realizado']).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Meta': f"R$ {meta:,.2f}".format(grupo_data['Realizado']).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção vs Meta': f"{100 * projecao / meta:.2f}%" if meta else "Meta não definida"
            })
        total_realizado += grupo_data['Realizado']
        total_projecao += grupo_data['Projeção']
        total_meta += grupo_data['Meta']

        grupo_data['Projeção vs Meta'] = f"{100 * grupo_data['Projeção'] / grupo_data['Meta']:.2f}%" if grupo_data['Meta'] else "Meta não definida"
        # Formata valores para o grupo
        grupo_data['Realizado'] = "R$ {:,.2f}".format(grupo_data['Realizado']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Projeção'] = "R$ {:,.2f}".format(grupo_data['Projeção']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Meta'] = "R$ {:,.2f}".format(grupo_data['Meta']).replace(',', 'X').replace('.', ',').replace('X', '.')

        # Insere os dados agregados do grupo no início de sua seção
        data.insert(len(data) - len(tipos), grupo_data)

    total_row = {
        'Subcategoria': 'Totais',
        'Realizado': f"R$ {total_realizado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção': f"R$ {total_projecao:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Meta': f"R$ {total_meta:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    }
    data.insert(0, total_row)  # Insere a linha de totais no início

    df_outros_agregados = df[~df['Tipo_Produto'].isin(sum(subcategorias.values(), [])) & (df['Grupo'] == 'AGREGADOS')]
    # Adapte as funções de cálculo se necessário para aceitar um df como argumento

    realizado_outros = calcular_somas_grupos(df_outros_agregados, "OUTROS AGREGADOS")
    projecao_outros = calc_projecao_categoria(realizado_outros)
    meta_outros = calcular_media_faturamento_ultimos_3_meses(df_outros_agregados, ["OUTROS AGREGADOS"])
    projecao_vs_meta_outros = f"{(projecao_outros / meta_outros) * 100:.2f}%" if meta_outros > 0 else " Meta não definida"

    realizado_frete = calcular_somas_grupos_frete(df_frete)
    projecao_frete = calc_projecao_categoria(realizado_frete)
    meta_frete = calcular_media_faturamento_ultimos_3_meses_frete(df_frete)
    projecao_vs_meta_frete = f"{(projecao_frete / meta_frete * 100) if meta_frete else 0:.2f}%"

    # Calcula dados para 'FRETE'
    realizado_frete = calcular_somas_grupos_frete(df_frete)
    projecao_frete = calc_projecao_categoria(realizado_frete)
    meta_frete = calcular_media_faturamento_ultimos_3_meses_frete(df_frete)
    projecao_vs_meta_frete = f"{(projecao_frete / meta_frete * 100) if meta_frete else 0:.2f}%"

    data.append({
        'Subcategoria': '· FRETE',
        'Realizado': f"R$ {realizado_frete:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção': f"R$ {projecao_frete:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Meta': f"R$ {meta_frete:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção vs Meta': projecao_vs_meta_frete
    })

    # Calcula dados para 'BENEFICIAMENTO'
    df_benef_filtered = df_benef[df_benef['NOME_BENEF'] != 'Caixa de Madeira']
    realizado_benef = calcular_somas_grupos_benef(df_benef_filtered)
    projecao_benef = calc_projecao_categoria(realizado_benef)
    meta_benef = calcular_media_faturamento_ultimos_3_meses_benef(df_benef_filtered)
    projecao_vs_meta_benef = f"{(projecao_benef / meta_benef * 100) if meta_benef else 0:.2f}%"

    data.append({
        'Subcategoria': '· BENEFICIAMENTO',
        'Realizado': f"R$ {realizado_benef:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção': f"R$ {projecao_benef:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Meta': f"R$ {meta_benef:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção vs Meta': projecao_vs_meta_benef
    })

    realizado_madeira = calcular_somas_grupos_benef(df_madeira)
    projecao_madeira = calc_projecao_categoria(realizado_madeira)
    meta_madeira = calcular_media_faturamento_ultimos_3_meses_benef(df_madeira)
    projecao_vs_meta_madeira = f"{(projecao_madeira / meta_madeira * 100) if meta_madeira else 0:.2f}%"

    data.append({
        'Subcategoria': '· CAIXA DE MADEIRA',
        'Realizado': f"R$ {realizado_madeira:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção': f"R$ {projecao_madeira:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Meta': f"R$ {meta_madeira:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção vs Meta': projecao_vs_meta_madeira
    })

    style_cell = {
    'textAlign': 'left',
    'padding': '0px',
    'whiteSpace': 'normal',
    }

    style_header = {
        'fontWeight': 'bold',
        'textAlign': 'left',
    }

    return dash_table.DataTable(
        id='categoria-agregadas-table',
        columns=[
            {'name': 'Subcategoria', 'id': 'Subcategoria'},
            {'name': 'Realizado', 'id': 'Realizado'},
            {'name': 'Projeção', 'id': 'Projeção'},
            {'name': 'Meta', 'id': 'Meta', 'type': 'numeric'},
            {'name': 'Projeção vs Meta', 'id': 'Projeção vs Meta'}
        ],
        data=data,
        style_table={'overflowX': 'auto'},
        style_cell=style_cell,
        style_header=style_header,
        style_data_conditional=[
            {'if': {'row_index': 0}, 'backgroundColor': '#F1F1F1', 'fontWeight': 'bold'},  
            # Estilos para os grupos
            {'if': {'filter_query': '{Subcategoria} = "KIT\'S PARA BOX"'},
             'backgroundColor': '#b0c4de', 'color': 'black'},
            {'if': {'filter_query': '{Subcategoria} = "KIT\'S PARA JANELA"'},
             'backgroundColor': '#b0c4de', 'color': 'black'},
            {'if': {'filter_query': '{Subcategoria} = "PERFIS PARA VIDRO TEMPERADO"'},
             'backgroundColor': '#b0c4de', 'color': 'black'},
            {'if': {'filter_query': '{Subcategoria} = "FERRAGENS"'},
             'backgroundColor': '#b0c4de', 'color': 'black'},
            {'if': {'filter_query': '{Subcategoria} = "SERVIÇOS"'},
             'backgroundColor': '#b0c4de', 'color': 'black'},
            {'if': {'filter_query': '{Subcategoria} = "OUTROS"'},
             'backgroundColor': '#b0c4de', 'color': 'black'}
        ])

def calcular_pontuacao(porcentagem):
    if porcentagem < 80:
        return 100
    elif porcentagem < 90:
        return 200
    elif porcentagem < 100:
        return 300
    else:
        return 400

# Função auxiliar para carregar imagens e converter para o formato adequado para uso no Dash
def encode_image(image_file):
    encoded = base64.b64encode(open(image_file, 'rb').read())
    return f'data:image/png;base64,{encoded.decode()}'

# Função para escolher a imagem com base na porcentagem
def image_for_percentage(percentage):
    if percentage >= 100:
        return encode_image('assets/img/01feliz.png')
    elif percentage >= 50:
        return encode_image('assets/img/02triste.png')
    else:
        return encode_image('assets/img/03triste.png')

@app.callback(
    Output("pontuacao-vendedor-destaque", "children"), 
    [Input("vendedor-dropdown", "value")]
)

def atualizar_pontuacao_vendedor(vendedor_selecionado):
    if vendedor_selecionado == 'TODOS OS VENDEDORES':
        return "Selecione um vendedor"

    temp_path = 'META_VENDEDORES.xlsx'
    df_meta_vendedores = pd.read_excel(temp_path)

    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y', errors='coerce')
    
    meta_vendedor = df_meta_vendedores[df_meta_vendedores['NOME VENDEDOR'] == vendedor_selecionado]
    if not meta_vendedor.empty:
        meta_vidro = meta_vendedor['META VIDRO'].values[0]
        meta_agregado = meta_vendedor['META AGREGADOS'].values[0]
        meta_geral = meta_vendedor['META VENDEDOR'].values[0]

        def projecao_vidro():
            df_ = df
            df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')
            df_filtrado = df_[(df_['Vendedor'] == vendedor_selecionado) & 
                            (df_['Data_Pedido'] >= start_date_realizado) & 
                            (df_['Data_Pedido'] <= end_date_realizado) & 
                            (df_['Grupo'] == "VIDRO")]
            df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')
            valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()
            dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
            total_dias_uteis = total_dias_uteis_no_mes(start_date_realizado)

            if dias_uteis_ate_ontem_ > 0:  
                projecao = (valor_realizado_vendedor / dias_uteis_ate_ontem_) * total_dias_uteis
            else: projecao = 0
            return projecao
        def projecao_agregado():
            agregados = ['ACESSÓRIOS', 'ALUMÍNIO', 'FERRAGEM', 'KIT PARA BOX PADRÃO', 'SILICONE']
            df_ = df
            df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')

            df_filtrado = df_[(df_['Vendedor'] == vendedor_selecionado) & 
                            (df_['Data_Pedido'] >= start_date_realizado) & 
                            (df_['Data_Pedido'] <= end_date_realizado) & 
                            (df_['Grupo'].isin(agregados))]
            
            df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')
            valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()
            dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
            total_dias_uteis = total_dias_uteis_no_mes(start_date_realizado)
           
            if dias_uteis_ate_ontem_ > 0: 
                projecao = (valor_realizado_vendedor / dias_uteis_ate_ontem_) * total_dias_uteis
            else: projecao = 0
            return projecao

        projecao_vidro_ = projecao_vidro()
        projecao_agregado_ = projecao_agregado()
        projecao_ = calc_projecao_vendedor(df, vendedor_selecionado)

        porcentagem_vidro = (projecao_vidro_ / meta_vidro) * 100 if meta_vidro > 0 else 0
        porcentagem_agregado = (projecao_agregado_ / meta_agregado) * 100 if meta_agregado > 0 else 0
        porcentagem_vendedor = (projecao_ / meta_geral) * 100 if meta_geral > 0 else 0
        
        pontuacao_vidro = calcular_pontuacao(porcentagem_vidro)
        pontuacao_agregado = calcular_pontuacao(porcentagem_agregado)
        pontuacao_geral = calcular_pontuacao(porcentagem_vendedor)

        image_vidro = image_for_percentage(porcentagem_vidro)
        image_agregado = image_for_percentage(porcentagem_agregado)
        image_geral = image_for_percentage(porcentagem_vendedor)

        return html.Div([
    html.Div([
        html.Div([
            html.Img(src=image_geral, style={'height': '60px', 'width': '60px'}),
            html.Div([
                html.P(f"GERAL {porcentagem_vendedor:.2f}%", style={'margin': 0, 'font-size': '1rem'}),
                html.P(f"{pontuacao_geral} PTS", style={'margin': '0 0 0 20px', 'font-size': '13px'}),  
            ], style={'display': 'flex', 'flex-direction': 'column', 'align-items': 'center'}),
        ], style={'display': 'flex', 'align-items': 'center', 'margin-right': 'auto'}),
        html.Div([
            html.Img(src=image_vidro, style={'height': '60px', 'width': '60px'}),
            html.Div([
                html.P(f"VIDRO {porcentagem_vidro:.2f}%", style={'margin': 0, 'font-size': '1rem'}),
                html.P(f"{pontuacao_vidro} PTS", style={'margin': 0, 'font-size': '13px'}),
            ], style={'display': 'flex', 'flex-direction': 'column', 'align-items': 'center'}),
        ], style={'display': 'flex', 'align-items': 'center'}),
        html.Div([
            html.Img(src=image_agregado, style={'height': '60px', 'width': '60px'}),
            html.Div([
                html.P(f"AGREGADO {porcentagem_agregado:.2f}%", style={'margin': '0 20px 0 0', 'font-size': '1rem'}),  
                html.P(f"{pontuacao_agregado} PTS", style={'margin': 0, 'font-size': '13px'}),
            ], style={'display': 'flex', 'flex-direction': 'column', 'align-items': 'center'}),
        ], style={'display': 'flex', 'align-items': 'center', 'margin-left': 'auto'}),  
    ], style={'display': 'flex', 'justify-content': 'space-between', 'width': '100%'}),  #
], style={'display': 'flex', 'flex-direction': 'column', 'align-items': 'center', 'justify-content': 'center', 'width': '100%'})

    else:
        print("Nenhum vendedor correspondente encontrado.") 

@app.callback(
    Output('meta-vendedor-texto', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_meta_vendedor(vendedor_selecionado):
    if vendedor_selecionado and vendedor_selecionado != 'TODOS OS VENDEDORES':
        meta_vendedor = df_metas[df_metas['NOME VENDEDOR'] == vendedor_selecionado]['META VENDEDOR'].values[0]
        return f"R$ {meta_vendedor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    else:
        return "Selecionar vendedor"

mes_atual_nome = datetime.now().strftime('%B').capitalize()
mensagem_atualizacao = f"Última atualização - Metas {mes_atual_nome}"
hoje_ = pd.to_datetime('today').normalize()
ontem_ = hoje_ - timedelta(days=1)
primeiro_dia_do_mes_ = hoje.replace(day=1)
ultimo_dia_do_mes_ = hoje + pd.offsets.MonthEnd(1)

dias_corridos_ = dias_uteis_ate_ontem(primeiro_dia_do_mes_, ontem_)
dias_uteis_mes_ = total_dias_uteis_no_mes(hoje_)
dias_restantes_ = dias_uteis_ate_ontem(hoje_, ultimo_dia_do_mes_)

########## LAYOUT DASH
app.layout = dbc.Container([
    dcc.Interval(
        id='interval-update', 
        interval=300*1000,  # 5 minutos
        n_intervals=0
    ),
    dcc.Store(id='meta-value-store'),
    dbc.Row([
        dbc.Col(html.H1("FAROL DE VENDAS", className="text-center-titulo"), width=12),
        html.P(mensagem_atualizacao, className="card-text", style={'color': '#3FB9C6', 'margin-top': '0px'})
    ]),
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardBody([html.H5("Filtro Vendedor", className="card-title", style={'text-align': 'left'}),
            dcc.Dropdown(
                id='vendedor-dropdown', 
                options=[{'label': 'TODOS OS VENDEDORES', 'value': 'TODOS OS VENDEDORES'}] + get_vendedor_names(df),
                value='TODOS OS VENDEDORES',
                clearable=False,
                 style={'width': '100%', 'border': 'none', 'background-color': 'transparent', 'font-weight': 'bold'}  # define a largura do dropdown
            )])]),
            width={"size": 3, "offset": 0}, className="mb-4")]),
    dbc.Row([
    dbc.Col(dbc.Card([dbc.CardBody([
    html.Div([
        html.Img(src=app.get_asset_url("img/iconmeta.svg"), style={'height': '50px', 'width': '50px'}),
        html.H5("META GERAL", className="card-title"),
    ], style={'display': 'flex', 'align-items': 'center'}),
    html.Div(id="meta-geral-texto", children=meta_geral_formatada, className="card-text", style={'fontSize': '1.2rem'}),  # Tamanho da fonte ajustado
    html.Div([
        html.P(f"Dias corridos: {dias_corridos_}", className="card-text", style={'font-size': '0.6rem', 'display': 'inline', 'margin-right': '10px', 'color': '#A3AED0'}),
        html.P(f"Dias úteis: {dias_uteis_mes_}", className="card-text", style={'font-size': '0.6rem', 'display': 'inline', 'margin-right': '10px', 'color': '#A3AED0'}),
        html.P(f"Dias restantes: {int(dias_restantes_)-1}", className="card-text", style={'font-size': '0.6rem', 'display': 'inline', 'color': '#A3AED0'})
    ], style={'display': 'flex', 'justify-content': 'center'})
], className="card-topo")]), width=3),
        dbc.Col(dbc.Card([dbc.CardBody([
                        html.Div([
                            html.Img(src=app.get_asset_url("img/iconfinanc.svg"), style={'height': '50px', 'width': '50px'}),
                            html.H5("REALIZADO GERAL", className="card-title"),
                              ], style={'display': 'flex', 'align-items': 'center'}),
                            html.P(f"R$ {valor_realizado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), className="card-text")
        ],className="card-topo")]), width=2),
        dbc.Col(
    dbc.Card([
        dbc.CardBody([
            html.Div([
                html.Img(
                    src=app.get_asset_url("img/iconproje.svg"),
                    style={'height': '50px', 'width': '50px'}
                ),
                html.H5("PROJEÇÃO GERAL", className="card-title"),
            ], style={'display': 'flex', 'align-items': 'center'}),
            html.P(
                [
                    f"R$ {valor_projetado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                    " - ", 
                    html.Span(
                        f"{percentual_comissao_str}",
                        style={'color': 'darkred'},
                        id='percentual-comissao'
                    ),
                ],
                className="card-text"
            ),
            dbc.Tooltip(
                tooltip_text,
                target='percentual-comissao',
                placement='top',
                is_open=False,
            )
        ], className="card-topo"),
    ]),
    width=2
),
        dbc.Col(dbc.Card([dbc.CardBody([
                         html.H5("PONTUAÇÃO VENDEDOR DESTAQUE", className="card-title", style={"text-align": "center", "margin-bottom": "1rem"}),
             dbc.Row([
                  html.Div(id="pontuacao-vendedor-destaque"),
            ],id='card-pontuacao-vendedor-destaque', justify="center", style={"margin-bottom": "0.5rem"}),
            
        ], className="card-topo")]), width=5)
    ], className="mb-4-1"),
        dbc.Row([
            dbc.Col(dbc.Card(dbc.CardBody([
            dcc.Graph(
                id='VENDAS POR CATEGORIA ÚTIMOS 3 MESES',
                figure=fig_pilha,
                style={'height': '100%', 'width': '100%', },  
                config={'responsive': True},className="graph-titulo")
            ]), 
                 style={'height': '380px', 'backgroundColor': 'white', 'margin-top': '10px'},
            ),width=4, className="mb-4"
        ),
       dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.Div([
            html.Img(src=app.get_asset_url("img/icometavend.svg"), style={'height': '30px', 'width': '30px', 'margin-right': '10px'}),
            html.H5("META POR VENDEDOR", className="card-title", style={'display': 'inline-block'}),
        ], style={'display': 'flex', 'align-items': 'center'}),
        html.Div([
            html.Div(id="meta-vendedor-texto", className="card-text", style={'marginTop': '10px'}),
        ], style={'marginTop': '10px'}),  
    ], style={'display': 'block'})
            ], className="mb-3"),
            dbc.Card([
                dbc.CardBody([
                    html.Div([
            html.Img(src=app.get_asset_url("img/icoreavend.svg"), style={'height': '30px', 'width': '30px', 'margin-right': '10px'}),
            html.H5("REALIZADO POR VENDEDOR", className="card-title", style={'display': 'inline-block'}),
        ], style={'display': 'flex', 'align-items': 'center'}),

        html.Div([
            html.P(id="realizado-vendedor", className="card-text"),
        ], style={'marginTop': '10px'}),  
    ])
            ], className="mb-3"),
            dbc.Card([
                dbc.CardBody([
                     html.Div([
            html.Img(src=app.get_asset_url("img/icoprojvend.svg"), style={'height': '30px', 'width': '30px', 'margin-right': '10px'}),
            html.H5("PROJEÇÃO POR VENDEDOR", className="card-title", style={'display': 'inline-block'}),
        ], style={'display': 'flex', 'align-items': 'center'}),
        html.Div([
            html.P(id="projecao-vendedor", className="card-text"),
        ], style={'marginTop': '10px'}),  
    ])
            ], className="mb-3"),
        ], width=12, md=6, lg=2, className="mb-3-1"),
    
    dbc.Col(dbc.Card(dbc.CardBody([
            dcc.Graph(
                        id='right-chart',
                        style={'height': '100%', 'width': '100%'},
                        config={'responsive': True}
                    )
                ]),
                style={'height': '380px', 'backgroundColor': 'white', 'margin-top': '10px'}, 
    )
        ),], className="mb-4"),
    
    dbc.Row([
    dbc.Card([
        dbc.CardBody([
            dbc.Row([
                # QUANTIDADE DE CLIENTES ATENDIDOS
                dbc.Col(dbc.Card([
                    dbc.CardBody([
                        dbc.Row([
                            dbc.Col(html.Img(
                                src=app.get_asset_url("img/client-atend.svg"), 
                                style={'height': '30px', 'width': '30px'}
                            ), width=12, className="text-center"),   
                        ]),
                        html.H6("QUANTIDADE DE CLIENTES ATENDIDOS", className="card-title", style={'margin-top': '20px'}),
                        dbc.Row([
                            dbc.Col(html.Div("AGREGADO", className="text-center", style={'fontSize': '14px'}), width=4),
                            dbc.Col(html.Div("COMUM", className="text-center", style={'fontSize': '14px'}), width=4),
                            dbc.Col(html.Div("TEMPERADO", className="text-center", style={'fontSize': '13px'}), width=4)
                        ]),
                        dbc.Row([
                            dbc.Col(html.Div(id="clientes-atendidos-vidro", className="text-center"), width=4),
                            dbc.Col(html.Div(id="clientes-atendidos-agregados", className="text-center"), width=4),
                            dbc.Col(html.Div(id="clientes-atendidos-temperados", className="text-center"), width=4)
                        ])
                    ])
                ], style={'backgroundColor': 'white', 'border-color': '#FFFFFF'}), width=3, md=3, sm=6, xs=12),

                # VENDA POR LOCALIDADE
                dbc.Col(dbc.Card([
                    dbc.CardBody([
                        dbc.Row([
                            dbc.Col(html.Img(
                                src=app.get_asset_url("img/localiza.svg"),
                                style={'height': '30px', 'width': '30px'}
                            ), width=12, className="text-center"),   
                        ]),
                        html.H6("VENDA POR LOCALIDADE", className="card-title" , style={'margin-top': '20px'}),
                        dbc.Row([
                            dbc.Col(html.Div("CAPITAL", className="text-center small-text", style={'fontSize': '14px'}), width=6),
                            dbc.Col(html.Div("INTERIOR", className="text-center small-text", style={'fontSize': '14px'}), width=6)
                        ]),
                        dbc.Row([
                            dbc.Col(html.Div(id="vendas-capital", className="text-center"), width=6),
                            dbc.Col(html.Div(id="vendas-interior", className="text-center"), width=6)
                        ])
                    ])
                ], style={'backgroundColor': 'white', 'border-color': '#FFFFFF'}), width=2, md=3, sm=6, xs=12),

                # RECOMPRA DOS ÚLTIMOS 6 MESES
                dbc.Col(dbc.Card([
                        dbc.CardBody(id="recompra-ultimos-6-meses", style={'backgroundColor': 'white', 'border-radius': '20px'})
                    ], style={'backgroundColor': 'white', 'border-color': '#FFFFFF'}), width=7, md=6, sm=12, xs=12),
            ])
        ])
    ], style={'backgroundColor': 'white', 'border-radius': '20px'})
], className="mb-4"),
    
dbc.Row([
    dbc.Col(
        html.Div(id='faturamento_vidro_card_container', className="card-style col-equal-height table table-container"), 
         style={'overflowX': 'auto'},
        width=6, 
        className="mb-4"
    ),
    dbc.Col(
        html.Div(id='categoria_vidro_table_container', className="card-style col-equal-height table table-container"), 
        style={'overflowX': 'auto'},
        width=6, 
        className="mb-4"
    ),
    ]),

dbc.Row([
    dbc.Col(
        html.Div(id='faturamento_agregados_3m_container', className="card-style col-equal-height table table-container"), 
        style={'overflowX': 'auto'},
        width=6, 
        className="mb-4"
    ),
    dbc.Col(
        html.Div(id='categoria_agregados_table_container', className="card-style col-equal-height table table-container"), 
        style={'overflowX': 'auto'},
        width=6, 
        className="mb-4"
    ),
    ]),

    # Tabela Cliente Sintético
    dbc.Row([
        dbc.Col(cliente_sintetico_card, className="card-style-2", width=12),
    ]),

    dcc.Download(id='download-excel'),
    
    ], className="container")

if __name__ == "__main__":
    app.run_server(host='', debug=False)
